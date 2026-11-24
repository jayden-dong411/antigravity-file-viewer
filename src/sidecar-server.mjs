import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import {
  browseDirectory,
  expandPath,
  loadOfficeFile,
  loadZipEntry,
  loadZipEntryFromBuffer,
  parseOfficeBuffer,
} from './parsers.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN_ROOT = path.resolve(__dirname, '..');
const HTML_PATH = path.join(PLUGIN_ROOT, 'dist', 'sidecar.html');
const DATA_DIR =
  process.env.OFFICE_VIEWER_DATA_DIR ||
  path.join(os.homedir(), '.gemini', 'antigravity', 'plugin_data', 'office-viewer');
const RECENT_PATH = path.join(DATA_DIR, 'recent.json');
const ACTIVE_PATH = path.join(DATA_DIR, 'active.json');

const PORT = Number(process.env.ANTIGRAVITY_SIDECAR_WEB_PORT || process.env.PORT || 0);
const MAX_BODY_BYTES = 70 * 1024 * 1024; // 70 MiB (base64 overhead for up to 50 MiB file)

function getShortcuts() {
  const homeDir = os.homedir();
  return [
    { label: 'Home', path: homeDir },
    { label: 'Desktop', path: path.join(homeDir, 'Desktop') },
    { label: 'Documents', path: path.join(homeDir, 'Documents') },
    { label: 'Downloads', path: path.join(homeDir, 'Downloads') },
  ];
}

async function readJsonSafe(filePath, fallback) {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8'));
  } catch {
    return fallback;
  }
}

async function addRecent(entry) {
  try {
    await fs.mkdir(DATA_DIR, { recursive: true });
    const list = await readJsonSafe(RECENT_PATH, []);
    const filtered = (Array.isArray(list) ? list : []).filter((i) => i.filePath !== entry.filePath);
    const next = [
      {
        filePath: entry.filePath,
        fileName: entry.fileName,
        ext: entry.ext,
        size: entry.size,
        openedAt: Date.now(),
      },
      ...filtered,
    ].slice(0, 15);
    await fs.writeFile(RECENT_PATH, JSON.stringify(next, null, 2), 'utf8');
    await fs.writeFile(
      ACTIVE_PATH,
      JSON.stringify({ filePath: entry.filePath, fileName: entry.fileName, updatedAt: Date.now() }, null, 2),
      'utf8',
    );
    return next;
  } catch {
    return [];
  }
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let total = 0;
    const chunks = [];
    req.on('data', (chunk) => {
      total += chunk.length;
      if (total > MAX_BODY_BYTES) {
        req.destroy();
        reject(new Error('Payload too large'));
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8');
        resolve(raw ? JSON.parse(raw) : {});
      } catch (e) {
        reject(new Error('Invalid JSON body'));
      }
    });
    req.on('error', reject);
  });
}

function sendJson(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

const STATIC_MIME = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.pdf': 'application/pdf',
  '.mp4': 'video/mp4',
  '.webm': 'video/webm',
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
};

const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Sidecar-Token');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url || '/', 'http://127.0.0.1');
  const pathname = url.pathname;

  try {
    if (req.method === 'GET' && (pathname === '/health' || pathname === '/healthz')) {
      sendJson(res, 200, { ok: true });
      return;
    }

    if (req.method === 'GET' && (pathname === '/' || pathname === '/index.html')) {
      const html = await fs.readFile(HTML_PATH, 'utf8');
      res.writeHead(200, {
        'Content-Type': 'text/html; charset=utf-8',
        'Cache-Control': 'no-store',
      });
      res.end(html);
      return;
    }

    if (req.method === 'GET' && pathname === '/api/raw') {
      const target = url.searchParams.get('file') || '';
      const resolved = expandPath(target);
      const stat = await fs.stat(resolved);
      if (!stat.isFile()) {
        sendJson(res, 404, { error: 'Not a file' });
        return;
      }
      const ext = path.extname(resolved).toLowerCase();
      const buf = await fs.readFile(resolved);
      res.writeHead(200, {
        'Content-Type': STATIC_MIME[ext] || 'application/octet-stream',
        'Content-Disposition': 'inline',
        'Cache-Control': 'no-store',
      });
      res.end(buf);
      return;
    }

    // Serve local HTML + sibling relative assets (CSS/JS/images) cleanly under /site/<base64url-dir>/<relpath>
    if (req.method === 'GET' && pathname.startsWith('/site/')) {
      const rest = pathname.slice('/site/'.length);
      const slashIdx = rest.indexOf('/');
      if (slashIdx <= 0) {
        sendJson(res, 400, { error: 'Invalid site path' });
        return;
      }
      const dirEncoded = rest.slice(0, slashIdx);
      const relPath = decodeURIComponent(rest.slice(slashIdx + 1));
      const baseDir = path.resolve(Buffer.from(dirEncoded, 'base64url').toString('utf8'));
      const fullPath = path.resolve(baseDir, relPath);
      const relCheck = path.relative(baseDir, fullPath);
      if (relCheck.startsWith('..') || path.isAbsolute(relCheck)) {
        sendJson(res, 403, { error: 'Path escapes base directory' });
        return;
      }
      const buf = await fs.readFile(fullPath);
      const ext = path.extname(fullPath).toLowerCase();
      res.writeHead(200, {
        'Content-Type': STATIC_MIME[ext] || 'application/octet-stream',
        'Cache-Control': 'no-store',
      });
      res.end(buf);
      return;
    }

    if (req.method === 'GET' && pathname === '/api/state') {
      const recent = await readJsonSafe(RECENT_PATH, []);
      const active = await readJsonSafe(ACTIVE_PATH, null);
      sendJson(res, 200, {
        active,
        recent: Array.isArray(recent) ? recent : [],
        shortcuts: getShortcuts(),
      });
      return;
    }

    if (req.method === 'POST' && pathname === '/api/load') {
      const body = await readJsonBody(req);
      const data = await loadOfficeFile(body.filePath, { maxRows: body.maxRows });
      const recent = await addRecent(data);
      const siteUrl =
        data.kind === 'html' && data.filePath
          ? `/site/${Buffer.from(path.dirname(data.filePath), 'utf8').toString('base64url')}/${encodeURIComponent(path.basename(data.filePath))}`
          : undefined;
      sendJson(res, 200, { ...data, siteUrl, recent });
      return;
    }

    if (req.method === 'POST' && pathname === '/api/upload') {
      const body = await readJsonBody(req);
      const buffer = Buffer.from(String(body.base64 || ''), 'base64');
      const fileName = String(body.fileName || 'document');
      const parsed = await parseOfficeBuffer(buffer, fileName, {
        maxRows: body.maxRows,
        includeBase64: true,
      });
      sendJson(res, 200, {
        filePath: '',
        fileName,
        ext: path.extname(fileName).toLowerCase(),
        size: buffer.length,
        ...parsed,
      });
      return;
    }

    if (req.method === 'POST' && pathname === '/api/zip-entry') {
      const body = await readJsonBody(req);
      const entry = body.base64
        ? await loadZipEntryFromBuffer(Buffer.from(String(body.base64), 'base64'), String(body.entryPath || ''))
        : await loadZipEntry(String(body.filePath || ''), String(body.entryPath || ''));
      sendJson(res, 200, entry);
      return;
    }

    if (req.method === 'POST' && pathname === '/api/browse') {
      const body = await readJsonBody(req);
      const listing = await browseDirectory(body.dirPath);
      sendJson(res, 200, listing);
      return;
    }

    sendJson(res, 404, { error: 'Not found' });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    sendJson(res, 400, { error: message });
  }
});

server.listen(PORT, '127.0.0.1', () => {
  const addr = server.address();
  const actualPort = typeof addr === 'object' && addr ? addr.port : PORT;
  console.log(`[office-viewer-sidecar] Listening on http://127.0.0.1:${actualPort}`);
});

for (const sig of ['SIGINT', 'SIGTERM']) {
  process.on(sig, () => {
    server.close(() => process.exit(0));
  });
}
