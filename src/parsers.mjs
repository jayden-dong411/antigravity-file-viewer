import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { marked } from 'marked';

export const SUPPORTED_EXTENSIONS = new Set([
  '.docx', '.dotx',
  '.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.tsv', '.ods',
  '.pptx', '.pptm',
  '.md', '.markdown',
]);

const MAX_FILE_BYTES = 50 * 1024 * 1024; // 50 MiB
const DEFAULT_MAX_ROWS = 1000;
const DEFAULT_MAX_COLS = 100;

export function expandPath(inputPath) {
  if (!inputPath || typeof inputPath !== 'string') return '';
  const trimmed = inputPath.trim();
  if (trimmed.startsWith('~/') || trimmed === '~') {
    return path.resolve(os.homedir(), trimmed.slice(2));
  }
  return path.resolve(trimmed);
}

function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)));
}

function mimeFromExt(fileName) {
  const ext = path.extname(fileName).toLowerCase();
  const map = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
  };
  return map[ext] || null;
}

async function parseDocx(buffer) {
  const result = await mammoth.convertToHtml(
    { buffer },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        const b64 = await image.read('base64');
        return { src: `data:${image.contentType};base64,${b64}` };
      }),
    },
  );
  return {
    kind: 'docx',
    html: result.value || '<p><em>Empty document</em></p>',
    warnings: (result.messages || []).filter((m) => m.type === 'warning').map((m) => m.message).slice(0, 10),
  };
}

function parseSpreadsheet(buffer, ext, maxRows = DEFAULT_MAX_ROWS) {
  const readOpts = { type: 'buffer', cellDates: true };
  if (ext === '.tsv') readOpts.FS = '\t';
  const wb = XLSX.read(buffer, readOpts);

  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const ref = ws['!ref'];
    if (!ref) {
      return { name, rows: [], merges: [], totalRows: 0, totalCols: 0, truncated: false };
    }
    const range = XLSX.utils.decode_range(ref);
    const totalRows = range.e.r - range.s.r + 1;
    const totalCols = range.e.c - range.s.c + 1;
    const rowLimit = Math.min(totalRows, maxRows);
    const colLimit = Math.min(totalCols, DEFAULT_MAX_COLS);

    const rows = [];
    for (let r = 0; r < rowLimit; r++) {
      const row = [];
      for (let c = 0; c < colLimit; c++) {
        const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c });
        const cell = ws[addr];
        if (!cell) {
          row.push('');
        } else if (cell.w !== undefined && cell.w !== null) {
          row.push(String(cell.w));
        } else if (cell.v instanceof Date) {
          row.push(cell.v.toISOString().slice(0, 10));
        } else if (cell.v !== undefined && cell.v !== null) {
          row.push(String(cell.v));
        } else {
          row.push('');
        }
      }
      rows.push(row);
    }

    const merges = (ws['!merges'] || [])
      .map((m) => ({
        sr: m.s.r - range.s.r,
        sc: m.s.c - range.s.c,
        er: m.e.r - range.s.r,
        ec: m.e.c - range.s.c,
      }))
      .filter((m) => m.sr >= 0 && m.sc >= 0 && m.sr < rowLimit && m.sc < colLimit);

    return {
      name,
      rows,
      merges,
      totalRows,
      totalCols,
      truncated: totalRows > rowLimit || totalCols > colLimit,
    };
  });

  return { kind: 'sheet', sheets };
}

function extractTextFromXml(xml) {
  const texts = [];
  const matches = xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g);
  for (const m of matches) {
    texts.push(decodeXmlEntities(m[1]));
  }
  return texts.join('');
}

async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideEntries = [];

  zip.forEach((relativePath) => {
    const match = relativePath.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (match) {
      slideEntries.push({ index: Number(match[1]), path: relativePath });
    }
  });
  slideEntries.sort((a, b) => a.index - b.index);

  const notes = [];
  for (const entry of slideEntries) {
    const notesFile = zip.file(`ppt/notesSlides/notesSlide${entry.index}.xml`);
    if (notesFile) {
      const notesXml = await notesFile.async('text');
      const noteParts = [];
      for (const pm of notesXml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)) {
        const t = extractTextFromXml(pm[0]).trim();
        if (t && t !== String(entry.index)) noteParts.push(t);
      }
      const text = noteParts.join('\n').trim();
      if (text) notes.push({ index: entry.index, text });
    }
  }

  return {
    kind: 'pptx',
    slideCount: slideEntries.length,
    pptxBase64: buffer.toString('base64'),
    notes,
  };
}

export async function parseOfficeBuffer(buffer, fileName, options = {}) {
  const ext = path.extname(fileName).toLowerCase();
  if (ext === '.docx' || ext === '.dotx') {
    return await parseDocx(buffer);
  }
  if (['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.tsv', '.ods'].includes(ext)) {
    return parseSpreadsheet(buffer, ext, options.maxRows || DEFAULT_MAX_ROWS);
  }
  if (ext === '.pptx' || ext === '.pptm') {
    return await parsePptx(buffer);
  }
  if (ext === '.md' || ext === '.markdown') {
    const html = await marked.parse(buffer.toString('utf8'));
    return { kind: 'markdown', html };
  }
  throw new Error(`Unsupported file format: ${ext || '(no extension)'}. Supported: .docx, .xlsx, .xls, .csv, .tsv, .ods, .pptx, .md`);
}

export async function loadOfficeFile(filePath, options = {}) {
  const resolved = expandPath(filePath);
  const stat = await fs.stat(resolved);
  if (!stat.isFile()) {
    throw new Error(`Path is not a file: ${resolved}`);
  }
  if (stat.size > MAX_FILE_BYTES) {
    throw new Error(`File exceeds 50 MiB limit (${(stat.size / 1024 / 1024).toFixed(1)} MiB)`);
  }
  const buffer = await fs.readFile(resolved);
  const fileName = path.basename(resolved);
  const parsed = await parseOfficeBuffer(buffer, fileName, options);
  return {
    filePath: resolved,
    fileName,
    ext: path.extname(fileName).toLowerCase(),
    size: stat.size,
    mtimeMs: stat.mtimeMs,
    ...parsed,
  };
}

export async function browseDirectory(dirPath) {
  const resolved = expandPath(dirPath || os.homedir());
  const entries = await fs.readdir(resolved, { withFileTypes: true });
  const dirs = [];
  const files = [];

  for (const entry of entries) {
    if (entry.name.startsWith('.')) continue;
    const full = path.join(resolved, entry.name);
    if (entry.isDirectory()) {
      dirs.push({ name: entry.name, path: full, isDir: true });
    } else if (entry.isFile()) {
      const ext = path.extname(entry.name).toLowerCase();
      if (SUPPORTED_EXTENSIONS.has(ext)) {
        try {
          const st = await fs.stat(full);
          files.push({ name: entry.name, path: full, isDir: false, ext, size: st.size, mtimeMs: st.mtimeMs });
        } catch {
          // Skip unreadable files
        }
      }
    }
  }

  dirs.sort((a, b) => a.name.localeCompare(b.name));
  files.sort((a, b) => a.name.localeCompare(b.name));

  return {
    dirPath: resolved,
    parentPath: path.dirname(resolved) !== resolved ? path.dirname(resolved) : null,
    entries: [...dirs, ...files],
  };
}
