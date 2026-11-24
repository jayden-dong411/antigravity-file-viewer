import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import {
  browseDirectory,
  loadOfficeFile,
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
      sendJson(res, 200, { ...data, recent });
      return;
    }

    if (req.method === 'POST' && pathname === '/api/upload') {
      const body = await readJsonBody(req);
      const buffer = Buffer.from(String(body.base64 || ''), 'base64');
      const fileName = String(body.fileName || 'document');
      const parsed = await parseOfficeBuffer(buffer, fileName, { maxRows: body.maxRows });
      sendJson(res, 200, {
        filePath: '',
        fileName,
        ext: path.extname(fileName).toLowerCase(),
        size: buffer.length,
        ...parsed,
      });
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
