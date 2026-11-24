import {
  App,
  applyDocumentTheme,
  applyHostStyleVariables,
} from '@modelcontextprotocol/ext-apps';
import { init as initPptxPreview } from 'pptx-preview';

type SheetMerge = { sr: number; sc: number; er: number; ec: number };
type SheetData = {
  name: string;
  rows: string[][];
  merges: SheetMerge[];
  totalRows: number;
  totalCols: number;
  truncated: boolean;
};

type SlideNote = { index: number; text: string };
type ZipEntry = {
  path: string;
  name: string;
  dir: string;
  isDir: boolean;
  ext: string;
  size: number;
  compressedSize: number;
  date: string;
};
type ZipPreview = {
  entryPath: string;
  previewType: 'text' | 'image' | 'binary';
  size: number;
  text?: string;
  dataUrl?: string;
  message?: string;
};

type RecentItem = {
  filePath: string;
  fileName: string;
  ext: string;
  size: number;
  openedAt: number;
};

type Shortcut = { label: string; path: string };

type NotebookOutput =
  | { type: 'stream'; name: string; text: string }
  | { type: 'error'; ename: string; evalue: string; traceback: string }
  | { type: 'image'; src: string }
  | { type: 'html'; html: string }
  | { type: 'text'; text: string };

type NotebookCell = {
  index: number;
  cellType: 'markdown' | 'code' | 'raw';
  executionCount?: number | null;
  source: string;
  html?: string;
  outputs?: NotebookOutput[];
};

type LoadedDoc = {
  filePath: string;
  fileName: string;
  ext: string;
  size: number;
  kind: 'docx' | 'sheet' | 'pptx' | 'markdown' | 'pdf' | 'html' | 'zip' | 'ipynb' | 'media' | 'image';
  html?: string;
  warnings?: string[];
  sheets?: SheetData[];
  pptxBase64?: string;
  slideCount?: number;
  notes?: SlideNote[];
  pdfBase64?: string;
  source?: string;
  siteUrl?: string;
  entries?: ZipEntry[];
  fileCount?: number;
  dirCount?: number;
  totalSize?: number;
  totalCompressed?: number;
  kernel?: string;
  language?: string;
  cellCount?: number;
  cells?: NotebookCell[];
  mediaType?: 'video' | 'audio';
  mime?: string;
  mediaBase64?: string;
  isSvg?: boolean;
  dataUrl?: string;
  recent?: RecentItem[];
};

const IS_SIDECAR = (window as unknown as { __FILE_VIEWER_MODE__?: string }).__FILE_VIEWER_MODE__ === 'sidecar';
const app = IS_SIDECAR ? null : new App({ name: 'File Viewer', version: '1.1.0' }, {});
let lastActiveUpdatedAt = 0;

async function callBackend(
  toolName: 'app_load_file' | 'app_parse_upload' | 'app_browse_directory' | 'app_get_recent' | 'app_load_zip_entry',
  args: Record<string, unknown>,
): Promise<{ isError?: boolean; structuredContent?: Record<string, unknown> }> {
  if (IS_SIDECAR) {
    const endpointMap: Record<typeof toolName, { method: string; url: string }> = {
      app_load_file: { method: 'POST', url: '/api/load' },
      app_parse_upload: { method: 'POST', url: '/api/upload' },
      app_browse_directory: { method: 'POST', url: '/api/browse' },
      app_get_recent: { method: 'GET', url: '/api/state' },
      app_load_zip_entry: { method: 'POST', url: '/api/zip-entry' },
    };
    const ep = endpointMap[toolName];
    const resp = await fetch(ep.url, {
      method: ep.method,
      headers: ep.method === 'POST' ? { 'Content-Type': 'application/json' } : undefined,
      body: ep.method === 'POST' ? JSON.stringify(args) : undefined,
    });
    const json = (await resp.json()) as Record<string, unknown>;
    return { isError: !resp.ok || Boolean(json.error), structuredContent: json };
  }
  const res = await app!.callServerTool({ name: toolName, arguments: args });
  return {
    isError: res.isError,
    structuredContent: (res.structuredContent || {}) as Record<string, unknown>,
  };
}

let pptxPreviewer: ReturnType<typeof initPptxPreview> | null = null;
let cachedPptxKey = '';
let cachedPptxBuf: ArrayBuffer | null = null;
let pptxMountSeq = 0;
let uploadedBlobUrl = '';
let uploadedZipBase64 = '';

function setBlobUrl(b64?: string, mime = 'application/pdf'): string {
  if (uploadedBlobUrl) {
    URL.revokeObjectURL(uploadedBlobUrl);
    uploadedBlobUrl = '';
  }
  if (!b64) return '';
  const buf = base64ToArrayBuffer(b64);
  uploadedBlobUrl = URL.createObjectURL(new Blob([buf], { type: mime }));
  return uploadedBlobUrl;
}

function base64ToArrayBuffer(b64: string): ArrayBuffer {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) {
    bytes[i] = bin.charCodeAt(i);
  }
  return bytes.buffer;
}

const state: {
  doc: LoadedDoc | null;
  activeSheet: number;
  zoom: number;
  search: string;
  showNotes: boolean;
  pptMode: 'list' | 'slide';
  htmlMode: 'preview' | 'source';
  htmlViewport: '100%' | '768px' | '390px';
  imgBg: 'checker' | 'light' | 'dark';
  imgRotate: number;
  imgDims: string;
  svgMode: 'preview' | 'source';
  playbackRate: number;
  zipPreview: ZipPreview | null;
  zipPreviewLoading: boolean;
  drawerOpen: boolean;
  currentDir: string;
  parentDir: string | null;
  dirEntries: { name: string; path: string; isDir: boolean; ext?: string; size?: number }[];
  recent: RecentItem[];
  shortcuts: Shortcut[];
  loading: boolean;
  error: string;
  selectedCell: { ref: string; value: string } | null;
  canUpdateContext: boolean;
} = {
  doc: null,
  activeSheet: 0,
  zoom: 100,
  search: '',
  showNotes: false,
  pptMode: 'list',
  htmlMode: 'preview',
  htmlViewport: '100%',
  imgBg: 'checker',
  imgRotate: 0,
  imgDims: '',
  svgMode: 'preview',
  playbackRate: 1,
  zipPreview: null,
  zipPreviewLoading: false,
  drawerOpen: false,
  currentDir: '',
  parentDir: null,
  dirEntries: [],
  recent: [],
  shortcuts: [],
  loading: false,
  error: '',
  selectedCell: null,
  canUpdateContext: false,
};

function esc(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fmtSize(bytes: number): string {
  if (!bytes || bytes < 1024) return `${bytes || 0} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function colLetter(idx: number): string {
  let n = idx + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function badgeClass(ext: string): string {
  const e = (ext || '').toLowerCase();
  if (e === '.docx' || e === '.dotx') return 'badge-word';
  if (['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.tsv', '.ods'].includes(e)) return 'badge-excel';
  if (e === '.pptx' || e === '.pptm') return 'badge-ppt';
  if (e === '.pdf') return 'badge-pdf';
  if (e === '.html' || e === '.htm') return 'badge-html';
  if (e === '.zip') return 'badge-zip';
  if (e === '.ipynb') return 'badge-ipynb';
  if (['.mp4', '.webm', '.mov', '.m4v', '.mp3', '.wav', '.m4a', '.ogg', '.flac'].includes(e)) return 'badge-media';
  if (['.svg', '.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.ico'].includes(e)) return 'badge-img';
  return 'badge-md';
}

function badgeLabel(ext: string): string {
  return (ext || '').replace(/^\./, '').toUpperCase() || 'DOC';
}

async function loadFileByPath(filePath: string) {
  if (!filePath.trim()) return;
  state.loading = true;
  state.error = '';
  state.selectedCell = null;
  render();
  try {
    const res = await callBackend('app_load_file', { filePath: filePath.trim() });
    const sc = res.structuredContent || {};
    if (res.isError || sc.error) {
      state.error = String(sc.error || 'Failed to load file');
    } else {
      state.doc = sc as unknown as LoadedDoc;
      state.activeSheet = 0;
      state.search = '';
      state.zipPreview = null;
      state.imgRotate = 0;
      state.imgDims = '';
      state.svgMode = 'preview';
      state.playbackRate = 1;
      uploadedZipBase64 = '';
      setBlobUrl(undefined);
      if (Array.isArray(sc.recent)) state.recent = sc.recent as RecentItem[];
      state.drawerOpen = false;
    }
  } catch (err) {
    state.error = err instanceof Error ? err.message : String(err);
  } finally {
    state.loading = false;
    render();
  }
}

async function handleFileUpload(file: File) {
  state.loading = true;
  state.error = '';
  state.selectedCell = null;
  render();
  try {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    const base64 = btoa(binary);
    const res = await callBackend('app_parse_upload', { fileName: file.name, base64 });
    const sc = res.structuredContent || {};
    if (res.isError || sc.error) {
      state.error = String(sc.error || 'Failed to parse file');
    } else {
      state.doc = sc as unknown as LoadedDoc;
      state.activeSheet = 0;
      state.search = '';
      state.zipPreview = null;
      state.imgRotate = 0;
      state.imgDims = '';
      state.svgMode = 'preview';
      state.playbackRate = 1;
      uploadedZipBase64 = state.doc.kind === 'zip' ? base64 : '';
      if (state.doc.kind === 'pdf' && state.doc.pdfBase64) setBlobUrl(state.doc.pdfBase64, 'application/pdf');
      else if (state.doc.kind === 'media' && state.doc.mediaBase64) setBlobUrl(state.doc.mediaBase64, state.doc.mime || 'video/mp4');
      else setBlobUrl(undefined);
      state.drawerOpen = false;
    }
  } catch (err) {
    state.error = err instanceof Error ? err.message : String(err);
  } finally {
    state.loading = false;
    render();
  }
}

async function browseDir(dirPath?: string) {
  try {
    const res = await callBackend('app_browse_directory', dirPath ? { dirPath } : {});
    const sc = res.structuredContent || {};
    if (!res.isError && sc.dirPath) {
      state.currentDir = String(sc.dirPath);
      state.parentDir = sc.parentPath ? String(sc.parentPath) : null;
      state.dirEntries = Array.isArray(sc.entries) ? (sc.entries as typeof state.dirEntries) : [];
      render();
    }
  } catch (err) {
    state.error = err instanceof Error ? err.message : String(err);
    render();
  }
}

function renderSheet(sheet: SheetData): string {
  if (!sheet || sheet.rows.length === 0) {
    return `<div class="empty-inline">This sheet is empty.</div>`;
  }

  const mergeMap = new Map<string, { colspan: number; rowspan: number }>();
  const hiddenCells = new Set<string>();
  for (const m of sheet.merges || []) {
    mergeMap.set(`${m.sr}:${m.sc}`, {
      colspan: m.ec - m.sc + 1,
      rowspan: m.er - m.sr + 1,
    });
    for (let r = m.sr; r <= m.er; r++) {
      for (let c = m.sc; c <= m.ec; c++) {
        if (r !== m.sr || c !== m.sc) hiddenCells.add(`${r}:${c}`);
      }
    }
  }

  const colCount = Math.max(...sheet.rows.map((r) => r.length), 1);
  const query = state.search.trim().toLowerCase();

  let headerCells = '<th class="row-num">#</th>';
  for (let c = 0; c < colCount; c++) {
    headerCells += `<th>${colLetter(c)}</th>`;
  }

  let bodyRows = '';
  let visibleCount = 0;
  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (query) {
      const hit = row.some((cell) => cell.toLowerCase().includes(query));
      if (!hit && r !== 0) continue;
    }
    visibleCount++;
    let tds = `<td class="row-num">${r + 1}</td>`;
    for (let c = 0; c < colCount; c++) {
      const key = `${r}:${c}`;
      if (hiddenCells.has(key)) continue;
      const val = row[c] ?? '';
      const span = mergeMap.get(key);
      const attrs = [
        `data-r="${r}"`,
        `data-c="${c}"`,
        span && span.colspan > 1 ? `colspan="${span.colspan}"` : '',
        span && span.rowspan > 1 ? `rowspan="${span.rowspan}"` : '',
        query && val.toLowerCase().includes(query) ? 'class="cell-match"' : '',
      ]
        .filter(Boolean)
        .join(' ');
      tds += `<td ${attrs}>${esc(val)}</td>`;
    }
    bodyRows += `<tr class="${r === 0 ? 'first-row' : ''}">${tds}</tr>`;
  }

  const formulaBar = state.selectedCell
    ? `<div class="formula-bar"><span class="cell-ref">${esc(state.selectedCell.ref)}</span><span class="cell-val">${esc(state.selectedCell.value)}</span></div>`
    : '';

  const meta = `<div class="sheet-meta">${visibleCount} / ${sheet.totalRows} rows · ${sheet.totalCols} cols${
    sheet.truncated ? ' (showing first 1,000 rows)' : ''
  }</div>`;

  return `${formulaBar}<div class="sheet-wrap" style="font-size:${(state.zoom / 100) * 13}px"><table class="excel-table"><thead><tr>${headerCells}</tr></thead><tbody>${bodyRows}</tbody></table></div>${meta}`;
}

function renderPptxContainer(doc: LoadedDoc): string {
  const notesHtml =
    state.showNotes && doc.notes && doc.notes.length > 0
      ? `<div class="pptx-notes-panel">
          <div class="section-title">Speaker Notes</div>
          ${doc.notes
            .map(
              (n) => `<div class="pptx-note-item"><strong>Slide ${n.index}:</strong> ${esc(n.text).replace(/\n/g, '<br>')}</div>`,
            )
            .join('')}
        </div>`
      : '';
  return `<div class="pptx-canvas"><div id="pptx-stage" class="pptx-stage"></div>${notesHtml}</div>`;
}

async function mountPptxPreview() {
  const seq = ++pptxMountSeq;
  if (pptxPreviewer) {
    try {
      pptxPreviewer.destroy();
    } catch {
      // Ignore cleanup errors
    }
    pptxPreviewer = null;
  }
  const doc = state.doc;
  if (!doc || doc.kind !== 'pptx' || !doc.pptxBase64) return;
  const stage = document.getElementById('pptx-stage');
  if (!stage) return;

  const key = `${doc.filePath || doc.fileName}:${doc.size}`;
  if (cachedPptxKey !== key || !cachedPptxBuf) {
    cachedPptxKey = key;
    cachedPptxBuf = base64ToArrayBuffer(doc.pptxBase64);
  }

  const viewport = document.querySelector('.viewport') as HTMLElement | null;
  const availW = Math.max((viewport?.clientWidth || 960) - 48, 420);
  const baseW = Math.min(availW, 960);
  const width = Math.round(baseW * (state.zoom / 100));
  const height = Math.round((width * 9) / 16);

  try {
    const previewer = initPptxPreview(stage, {
      width,
      height: state.pptMode === 'slide' ? height : undefined,
      mode: state.pptMode,
    });
    pptxPreviewer = previewer;
    await previewer.preview(cachedPptxBuf.slice(0));
    if (seq !== pptxMountSeq) {
      try {
        previewer.destroy();
      } catch {
        // Stale mount
      }
    }
  } catch (err) {
    if (seq === pptxMountSeq) {
      stage.innerHTML = `<div class="banner error">PPTX render error: ${esc(err instanceof Error ? err.message : String(err))}</div>`;
    }
  }
}

async function openZipEntry(entryPath: string) {
  const doc = state.doc;
  if (!doc || doc.kind !== 'zip') return;
  state.zipPreviewLoading = true;
  render();
  try {
    const args: Record<string, unknown> = { entryPath };
    if (doc.filePath) args.filePath = doc.filePath;
    else if (uploadedZipBase64) args.base64 = uploadedZipBase64;
    const res = await callBackend('app_load_zip_entry', args);
    const sc = res.structuredContent || {};
    if (res.isError || sc.error) {
      state.zipPreview = {
        entryPath,
        previewType: 'binary',
        size: 0,
        message: String(sc.error || 'Failed to preview entry'),
      };
    } else {
      state.zipPreview = sc as unknown as ZipPreview;
    }
  } catch (err) {
    state.zipPreview = {
      entryPath,
      previewType: 'binary',
      size: 0,
      message: err instanceof Error ? err.message : String(err),
    };
  } finally {
    state.zipPreviewLoading = false;
    render();
  }
}

function renderPdf(doc: LoadedDoc): string {
  if (!IS_SIDECAR) {
    return `<div class="empty-inline">
      <p><strong>PDF preview requires the File Viewer Sidecar panel.</strong></p>
      <p class="muted">MCP App iframes enforce a strict CSP that blocks nested PDF viewers. Open this file in the File Viewer sidecar tab for native PDF viewing.</p>
    </div>`;
  }
  const src = doc.filePath
    ? `/api/raw?file=${encodeURIComponent(doc.filePath)}`
    : uploadedBlobUrl;
  if (!src) return `<div class="empty-inline">Unable to load PDF stream.</div>`;
  return `<div class="pdf-stage"><iframe class="pdf-frame" src="${esc(src)}" title="${esc(doc.fileName)}"></iframe></div>`;
}

function renderHtmlDoc(doc: LoadedDoc): string {
  if (state.htmlMode === 'source' || !IS_SIDECAR) {
    return `<div class="html-source-wrap"><pre class="code-view"><code>${esc(doc.source || '')}</code></pre></div>`;
  }
  const widthStyle = state.htmlViewport === '100%' ? 'width:100%' : `width:${state.htmlViewport}`;
  if (doc.siteUrl) {
    return `<div class="html-stage"><iframe class="html-frame" style="${widthStyle}" src="${esc(doc.siteUrl)}" sandbox="allow-scripts allow-same-origin allow-forms allow-popups" title="${esc(doc.fileName)}"></iframe></div>`;
  }
  return `<div class="html-stage"><iframe class="html-frame" style="${widthStyle}" srcdoc="${esc(doc.source || '')}" sandbox="allow-scripts allow-same-origin" title="${esc(doc.fileName)}"></iframe></div>`;
}

function renderZip(doc: LoadedDoc): string {
  const entries = doc.entries || [];
  const query = state.search.trim().toLowerCase();
  const filtered = query
    ? entries.filter((e) => e.path.toLowerCase().includes(query))
    : entries;

  const listHtml = filtered
    .map((e) => {
      if (e.isDir) {
        return `<div class="zip-item dir"><span class="icon-dir">DIR</span><span class="zip-path">${esc(e.path)}</span></div>`;
      }
      const active = state.zipPreview?.entryPath === e.path ? 'active' : '';
      return `<button class="zip-item file ${active}" data-zip-entry="${esc(e.path)}">
        <span class="badge ${badgeClass(e.ext)}">${badgeLabel(e.ext)}</span>
        <span class="zip-path" title="${esc(e.path)}">${esc(e.path)}</span>
        <span class="file-size">${fmtSize(e.size)}</span>
      </button>`;
    })
    .join('');

  let previewHtml = `<div class="empty-inline">Click any file in the archive to preview its contents.</div>`;
  if (state.zipPreviewLoading) {
    previewHtml = `<div class="empty-inline">Loading archive entry...</div>`;
  } else if (state.zipPreview) {
    const p = state.zipPreview;
    if (p.previewType === 'image' && p.dataUrl) {
      previewHtml = `<div class="zip-preview-head"><strong>${esc(p.entryPath)}</strong><span class="muted">${fmtSize(p.size)}</span></div>
        <div class="zip-img-wrap"><img src="${p.dataUrl}" alt="${esc(p.entryPath)}" /></div>`;
    } else if (p.previewType === 'text') {
      previewHtml = `<div class="zip-preview-head"><strong>${esc(p.entryPath)}</strong><span class="muted">${fmtSize(p.size)}</span></div>
        <pre class="code-view"><code>${esc(p.text || '')}</code></pre>`;
    } else {
      previewHtml = `<div class="zip-preview-head"><strong>${esc(p.entryPath)}</strong></div>
        <div class="empty-inline">${esc(p.message || 'Binary file')}</div>`;
    }
  }

  const ratio = doc.totalSize ? Math.round((1 - (doc.totalCompressed || 0) / doc.totalSize) * 100) : 0;
  return `
    <div class="zip-layout">
      <div class="zip-pane">
        <div class="zip-summary">${doc.fileCount || 0} files · ${doc.dirCount || 0} folders · ${fmtSize(doc.totalSize || 0)} (${ratio}% saved)</div>
        <div class="zip-list">${listHtml || '<div class="empty-inline">No matching entries.</div>'}</div>
      </div>
      <div class="zip-preview">${previewHtml}</div>
    </div>`;
}

function renderIpynb(doc: LoadedDoc): string {
  const cells = doc.cells || [];
  const query = state.search.trim().toLowerCase();
  const filtered = query
    ? cells.filter((c) => c.source.toLowerCase().includes(query))
    : cells;

  const cellsHtml = filtered
    .map((cell) => {
      if (cell.cellType === 'markdown') {
        return `<div class="nb-cell nb-md"><div class="nb-gutter">MD</div><div class="nb-md-body">${cell.html || ''}</div></div>`;
      }
      if (cell.cellType === 'raw') {
        return `<div class="nb-cell nb-raw"><div class="nb-gutter">RAW</div><pre class="nb-code"><code>${esc(cell.source)}</code></pre></div>`;
      }
      const prompt = cell.executionCount != null ? `In [${cell.executionCount}]:` : 'In [ ]:';
      const outputsHtml = (cell.outputs || [])
        .map((out) => {
          if (out.type === 'image') {
            return `<div class="nb-out nb-out-img"><img src="${out.src}" alt="Notebook output" /></div>`;
          }
          if (out.type === 'html') {
            return `<div class="nb-out nb-out-html">${out.html}</div>`;
          }
          if (out.type === 'error') {
            return `<pre class="nb-out nb-out-err"><code>${esc(out.traceback)}</code></pre>`;
          }
          if (out.type === 'stream') {
            return `<pre class="nb-out nb-out-stream ${out.name === 'stderr' ? 'stderr' : ''}"><code>${esc(out.text)}</code></pre>`;
          }
          return `<pre class="nb-out nb-out-text"><code>${esc(out.text)}</code></pre>`;
        })
        .join('');

      return `
        <div class="nb-cell nb-code-cell">
          <div class="nb-code-row">
            <div class="nb-gutter">${esc(prompt)}</div>
            <pre class="nb-code"><code>${esc(cell.source)}</code></pre>
          </div>
          ${outputsHtml ? `<div class="nb-outputs">${outputsHtml}</div>` : ''}
        </div>`;
    })
    .join('');

  return `
    <div class="nb-canvas" style="font-size:${(state.zoom / 100) * 13}px">
      <div class="nb-header">
        <strong>${esc(doc.kernel || 'Jupyter Notebook')}</strong>
        <span class="muted">${filtered.length} / ${doc.cellCount || cells.length} cells · ${esc(doc.language || 'python')}</span>
      </div>
      <div class="nb-cells">${cellsHtml || '<div class="empty-inline">No matching notebook cells.</div>'}</div>
    </div>`;
}

function renderMedia(doc: LoadedDoc): string {
  const src =
    IS_SIDECAR && doc.filePath
      ? `/api/raw?file=${encodeURIComponent(doc.filePath)}`
      : uploadedBlobUrl;

  if (!src) {
    return `<div class="empty-inline">
      <p><strong>Media streaming requires the File Viewer Sidecar panel.</strong></p>
      <p class="muted">Open this file in the File Viewer sidecar tab for native audio/video playback with seeking.</p>
    </div>`;
  }

  const player =
    doc.mediaType === 'audio'
      ? `<div class="audio-card">
          <div class="audio-title">${esc(doc.fileName)}</div>
          <div class="muted">${fmtSize(doc.size)} · ${esc(doc.mime || 'audio')}</div>
          <audio id="media-el" class="audio-player" controls src="${esc(src)}"></audio>
        </div>`
      : `<video id="media-el" class="video-player" controls playsinline src="${esc(src)}"></video>`;

  return `<div class="media-stage">${player}</div>`;
}

function renderImage(doc: LoadedDoc): string {
  if (doc.isSvg && state.svgMode === 'source') {
    return `<div class="html-source-wrap"><pre class="code-view"><code>${esc(doc.source || '')}</code></pre></div>`;
  }
  const src =
    IS_SIDECAR && doc.filePath && !doc.isSvg
      ? `/api/raw?file=${encodeURIComponent(doc.filePath)}`
      : doc.dataUrl || (IS_SIDECAR && doc.filePath ? `/api/raw?file=${encodeURIComponent(doc.filePath)}` : '');

  if (!src) {
    return `<div class="empty-inline">Unable to load image.</div>`;
  }

  const transform = `scale(${state.zoom / 100}) rotate(${state.imgRotate}deg)`;
  return `
    <div class="img-stage bg-${state.imgBg}">
      <img id="preview-img" class="preview-img" src="${esc(src)}" alt="${esc(doc.fileName)}" style="transform:${transform}" />
    </div>
    <div class="sheet-meta">${esc(doc.mime || 'image')} · ${fmtSize(doc.size)}${state.imgDims ? ` · ${esc(state.imgDims)}` : ''} · ${state.zoom}%${state.imgRotate ? ` · ${state.imgRotate}°` : ''}</div>`;
}

function renderDocument(): string {
  const doc = state.doc;
  if (!doc) {
    const recentHtml =
      state.recent.length > 0
        ? `<div class="section">
            <div class="section-title">Recent Files</div>
            <div class="file-list">
              ${state.recent
                .map(
                  (r) => `<button class="file-item" data-open="${esc(r.filePath)}">
                    <span class="badge ${badgeClass(r.ext)}">${badgeLabel(r.ext)}</span>
                    <span class="file-info">
                      <span class="file-name">${esc(r.fileName)}</span>
                      <span class="file-path">${esc(r.filePath)}</span>
                    </span>
                    <span class="file-size">${fmtSize(r.size)}</span>
                  </button>`,
                )
                .join('')}
            </div>
          </div>`
        : '';

    const shortcutsHtml =
      state.shortcuts.length > 0
        ? `<div class="shortcuts">
            ${state.shortcuts
              .map((s) => `<button class="chip" data-browse="${esc(s.path)}">${esc(s.label)}</button>`)
              .join('')}
          </div>`
        : '';

    return `
      <div class="welcome">
        <div class="dropzone" id="dropzone">
          <div class="drop-title">File Viewer</div>
          <div class="drop-sub">Drop a local <code>.docx</code>, <code>.xlsx</code>, <code>.pptx</code>, <code>.pdf</code>, <code>.html</code>, <code>.zip</code>, <code>.ipynb</code>, audio/video, image/SVG, or <code>.md</code> file here.</div>
          <div class="drop-actions">
            <button class="btn primary" id="pick-file-btn">Choose File</button>
            <button class="btn" id="open-drawer-btn">Browse Folders</button>
          </div>
          ${shortcutsHtml}
        </div>
        ${recentHtml}
      </div>`;
  }

  if (doc.kind === 'docx' || doc.kind === 'markdown') {
    const html = doc.html || '';
    const highlighted = state.search.trim()
      ? html
      : html;
    return `
      <div class="doc-canvas">
        <div class="paper" style="font-size:${(state.zoom / 100) * 14}px">
          ${highlighted}
        </div>
      </div>`;
  }

  if (doc.kind === 'sheet') {
    const sheets = doc.sheets || [];
    const current = sheets[state.activeSheet] || sheets[0];
    const tabs =
      sheets.length > 1
        ? `<div class="sheet-tabs">${sheets
            .map(
              (s, i) =>
                `<button class="sheet-tab ${i === state.activeSheet ? 'active' : ''}" data-sheet="${i}">${esc(s.name)}</button>`,
            )
            .join('')}</div>`
        : '';
    return `<div class="sheet-container">${tabs}${renderSheet(current)}</div>`;
  }

  if (doc.kind === 'pptx') {
    return renderPptxContainer(doc);
  }

  if (doc.kind === 'pdf') {
    return renderPdf(doc);
  }

  if (doc.kind === 'html') {
    return renderHtmlDoc(doc);
  }

  if (doc.kind === 'zip') {
    return renderZip(doc);
  }

  if (doc.kind === 'ipynb') {
    return renderIpynb(doc);
  }

  if (doc.kind === 'media') {
    return renderMedia(doc);
  }

  if (doc.kind === 'image') {
    return renderImage(doc);
  }

  return `<div class="empty-inline">Unsupported document view.</div>`;
}

function renderDrawer(): string {
  if (!state.drawerOpen) return '';
  const items = state.dirEntries
    .map((e) => {
      if (e.isDir) {
        return `<button class="drawer-item" data-dir="${esc(e.path)}">
          <span class="icon-dir">DIR</span>
          <span class="drawer-name">${esc(e.name)}</span>
        </button>`;
      }
      return `<button class="drawer-item" data-open="${esc(e.path)}">
        <span class="badge ${badgeClass(e.ext || '')}">${badgeLabel(e.ext || '')}</span>
        <span class="drawer-name">${esc(e.name)}</span>
        <span class="file-size">${fmtSize(e.size || 0)}</span>
      </button>`;
    })
    .join('');

  return `
    <aside class="drawer">
      <div class="drawer-head">
        <div class="drawer-path" title="${esc(state.currentDir)}">${esc(state.currentDir || '~')}</div>
        <div class="drawer-actions">
          ${state.parentDir ? `<button class="btn sm" data-dir="${esc(state.parentDir)}">Up</button>` : ''}
          <button class="btn sm" id="close-drawer-btn">Close</button>
        </div>
      </div>
      <div class="drawer-list">${items || '<div class="empty-inline">No supported Office files in this folder.</div>'}</div>
    </aside>`;
}

function render() {
  const root = document.getElementById('app');
  if (!root) return;

  const doc = state.doc;
  const pathVal = doc?.filePath || '';

  root.innerHTML = `
    <header class="toolbar">
      <div class="toolbar-row">
        <button class="btn sm" id="toggle-drawer-btn" title="Browse local folders">Browse</button>
        <form class="path-form" id="path-form">
          <input id="path-input" type="text" value="${esc(pathVal)}" placeholder="Enter local file path (e.g. ~/Documents/report.docx)..." />
          <button class="btn sm primary" type="submit">Open</button>
        </form>
        <input id="file-input" type="file" accept=".docx,.dotx,.xlsx,.xls,.xlsm,.xlsb,.csv,.tsv,.ods,.pptx,.pptm,.md,.markdown,.pdf,.html,.htm,.zip,.ipynb,.mp4,.webm,.mov,.m4v,.mp3,.wav,.m4a,.ogg,.flac,.svg,.png,.jpg,.jpeg,.webp,.gif,.bmp,.ico" hidden />
        <button class="btn sm" id="upload-btn" title="Choose a local file">Upload</button>
        ${doc?.filePath ? `<button class="btn sm" id="reload-btn" title="Reload from disk">Reload</button>` : ''}
      </div>
      ${
        doc
          ? `<div class="toolbar-row sub">
              <div class="doc-meta">
                <span class="badge ${badgeClass(doc.ext)}">${badgeLabel(doc.ext)}</span>
                <strong class="doc-title" title="${esc(doc.filePath || doc.fileName)}">${esc(doc.fileName)}</strong>
                <span class="muted">${fmtSize(doc.size)}</span>
              </div>
              <div class="controls">
                ${doc.kind !== 'pptx' && doc.kind !== 'pdf' && doc.kind !== 'html' && doc.kind !== 'media' && doc.kind !== 'image' ? `<input id="search-input" class="search-input" type="search" value="${esc(state.search)}" placeholder="Filter / search..." />` : ''}
                ${
                  doc.kind === 'pptx'
                    ? `<button class="btn sm" id="ppt-mode-btn" title="Toggle continuous list / single slide mode">${
                        state.pptMode === 'list' ? 'Continuous' : 'Single Slide'
                      }</button>`
                    : ''
                }
                ${
                  doc.kind === 'html'
                    ? `<button class="btn sm ${state.htmlMode === 'preview' ? 'active' : ''}" id="html-preview-btn">Preview</button>
                       <button class="btn sm ${state.htmlMode === 'source' ? 'active' : ''}" id="html-source-btn">Source</button>
                       ${
                         state.htmlMode === 'preview'
                           ? `<button class="btn sm ${state.htmlViewport === '100%' ? 'active' : ''}" data-vp="100%">Desktop</button>
                              <button class="btn sm ${state.htmlViewport === '768px' ? 'active' : ''}" data-vp="768px">Tablet</button>
                              <button class="btn sm ${state.htmlViewport === '390px' ? 'active' : ''}" data-vp="390px">Mobile</button>`
                           : ''
                       }`
                    : ''
                }
                ${
                  doc.kind === 'media'
                    ? [0.5, 1, 1.25, 1.5, 2]
                        .map(
                          (r) =>
                            `<button class="btn sm ${state.playbackRate === r ? 'active' : ''}" data-rate="${r}">${r}x</button>`,
                        )
                        .join('')
                    : ''
                }
                ${
                  doc.kind === 'image'
                    ? `${
                        doc.isSvg
                          ? `<button class="btn sm ${state.svgMode === 'preview' ? 'active' : ''}" id="svg-preview-btn">Preview</button>
                             <button class="btn sm ${state.svgMode === 'source' ? 'active' : ''}" id="svg-source-btn">Source</button>`
                          : ''
                      }
                      ${
                        !doc.isSvg || state.svgMode === 'preview'
                          ? `<button class="btn sm ${state.imgBg === 'checker' ? 'active' : ''}" data-imgbg="checker">Checker</button>
                             <button class="btn sm ${state.imgBg === 'light' ? 'active' : ''}" data-imgbg="light">Light</button>
                             <button class="btn sm ${state.imgBg === 'dark' ? 'active' : ''}" data-imgbg="dark">Dark</button>
                             <button class="btn sm" id="img-rotate-btn">Rotate</button>`
                          : ''
                      }`
                    : ''
                }
                ${doc.kind === 'pptx' && doc.notes && doc.notes.length > 0 ? `<button class="btn sm ${state.showNotes ? 'active' : ''}" id="notes-btn">Notes</button>` : ''}
                ${
                  doc.kind !== 'pdf' && doc.kind !== 'html' && doc.kind !== 'zip' && doc.kind !== 'media'
                    ? `<button class="btn sm" id="zoom-out">-</button>
                       <span class="zoom-label">${state.zoom}%</span>
                       <button class="btn sm" id="zoom-in">+</button>`
                    : ''
                }
                ${
                  state.canUpdateContext && doc.filePath
                    ? `<button class="btn sm" id="attach-btn" title="Attach file path to your next chat message">Attach</button>`
                    : ''
                }
                <button class="btn sm" id="home-btn">Home</button>
              </div>
            </div>`
          : ''
      }
    </header>
    ${state.error ? `<div class="banner error">${esc(state.error)}</div>` : ''}
    ${state.loading ? `<div class="banner info">Loading document...</div>` : ''}
    <main class="workspace">
      ${renderDrawer()}
      <section class="viewport">${renderDocument()}</section>
    </main>
  `;

  bindEvents();
  void mountPptxPreview();
}

function bindEvents() {
  const pathForm = document.getElementById('path-form');
  const pathInput = document.getElementById('path-input') as HTMLInputElement | null;
  pathForm?.addEventListener('submit', (e) => {
    e.preventDefault();
    if (pathInput?.value) void loadFileByPath(pathInput.value);
  });

  const fileInput = document.getElementById('file-input') as HTMLInputElement | null;
  const triggerPick = () => fileInput?.click();
  document.getElementById('upload-btn')?.addEventListener('click', triggerPick);
  document.getElementById('pick-file-btn')?.addEventListener('click', triggerPick);
  fileInput?.addEventListener('change', () => {
    const f = fileInput.files?.[0];
    if (f) void handleFileUpload(f);
  });

  const openDrawer = () => {
    state.drawerOpen = !state.drawerOpen;
    if (state.drawerOpen && !state.currentDir) {
      const startDir = state.doc?.filePath
        ? state.doc.filePath.replace(/\/[^/]+$/, '')
        : state.shortcuts[0]?.path;
      void browseDir(startDir);
    } else {
      render();
    }
  };
  document.getElementById('toggle-drawer-btn')?.addEventListener('click', openDrawer);
  document.getElementById('open-drawer-btn')?.addEventListener('click', () => {
    state.drawerOpen = true;
    void browseDir(state.shortcuts[0]?.path);
  });
  document.getElementById('close-drawer-btn')?.addEventListener('click', () => {
    state.drawerOpen = false;
    render();
  });

  document.getElementById('reload-btn')?.addEventListener('click', () => {
    if (state.doc?.filePath) void loadFileByPath(state.doc.filePath);
  });

  document.getElementById('home-btn')?.addEventListener('click', () => {
    state.doc = null;
    state.error = '';
    state.selectedCell = null;
    render();
  });

  document.getElementById('zoom-out')?.addEventListener('click', () => {
    state.zoom = Math.max(60, state.zoom - 10);
    render();
  });
  document.getElementById('zoom-in')?.addEventListener('click', () => {
    state.zoom = Math.min(200, state.zoom + 10);
    render();
  });

  document.getElementById('notes-btn')?.addEventListener('click', () => {
    state.showNotes = !state.showNotes;
    render();
  });

  document.getElementById('ppt-mode-btn')?.addEventListener('click', () => {
    state.pptMode = state.pptMode === 'list' ? 'slide' : 'list';
    render();
  });

  document.getElementById('html-preview-btn')?.addEventListener('click', () => {
    state.htmlMode = 'preview';
    render();
  });
  document.getElementById('html-source-btn')?.addEventListener('click', () => {
    state.htmlMode = 'source';
    render();
  });
  document.querySelectorAll<HTMLElement>('[data-vp]').forEach((el) => {
    el.addEventListener('click', () => {
      const vp = el.getAttribute('data-vp') as '100%' | '768px' | '390px';
      if (vp) {
        state.htmlViewport = vp;
        render();
      }
    });
  });

  document.querySelectorAll<HTMLElement>('[data-zip-entry]').forEach((el) => {
    el.addEventListener('click', () => {
      const p = el.getAttribute('data-zip-entry');
      if (p) void openZipEntry(p);
    });
  });

  // Media playback rate
  const mediaEl = document.getElementById('media-el') as HTMLMediaElement | null;
  if (mediaEl) {
    mediaEl.playbackRate = state.playbackRate;
  }
  document.querySelectorAll<HTMLElement>('[data-rate]').forEach((el) => {
    el.addEventListener('click', () => {
      const rate = Number(el.getAttribute('data-rate') || 1);
      state.playbackRate = rate;
      const m = document.getElementById('media-el') as HTMLMediaElement | null;
      if (m) m.playbackRate = rate;
      document.querySelectorAll<HTMLElement>('[data-rate]').forEach((b) => {
        b.classList.toggle('active', Number(b.getAttribute('data-rate')) === rate);
      });
    });
  });

  // Image & SVG controls
  document.getElementById('svg-preview-btn')?.addEventListener('click', () => {
    state.svgMode = 'preview';
    render();
  });
  document.getElementById('svg-source-btn')?.addEventListener('click', () => {
    state.svgMode = 'source';
    render();
  });
  document.querySelectorAll<HTMLElement>('[data-imgbg]').forEach((el) => {
    el.addEventListener('click', () => {
      const bg = el.getAttribute('data-imgbg') as 'checker' | 'light' | 'dark';
      if (bg) {
        state.imgBg = bg;
        render();
      }
    });
  });
  document.getElementById('img-rotate-btn')?.addEventListener('click', () => {
    state.imgRotate = (state.imgRotate + 90) % 360;
    render();
  });
  const previewImg = document.getElementById('preview-img') as HTMLImageElement | null;
  if (previewImg && !state.imgDims) {
    const updateDims = () => {
      if (previewImg.naturalWidth && previewImg.naturalHeight) {
        const dims = `${previewImg.naturalWidth} × ${previewImg.naturalHeight} px`;
        if (state.imgDims !== dims) {
          state.imgDims = dims;
          const meta = document.querySelector('.sheet-meta');
          if (meta && !meta.textContent?.includes('px')) {
            meta.textContent = `${meta.textContent} · ${dims}`;
          }
        }
      }
    };
    if (previewImg.complete) updateDims();
    else previewImg.addEventListener('load', updateDims, { once: true });
  }

  document.getElementById('attach-btn')?.addEventListener('click', async () => {
    if (!state.doc?.filePath || !app) return;
    try {
      await app.updateModelContext({
        content: [{ type: 'text', text: `Currently previewing Office document: ${state.doc.filePath}` }],
      });
    } catch {
      // Ignore if host surface does not support context attachment
    }
  });

  const searchInput = document.getElementById('search-input') as HTMLInputElement | null;
  searchInput?.addEventListener('input', () => {
    const pos = searchInput.selectionStart ?? searchInput.value.length;
    state.search = searchInput.value;
    render();
    const nextInput = document.getElementById('search-input') as HTMLInputElement | null;
    if (nextInput) {
      nextInput.focus();
      nextInput.setSelectionRange(pos, pos);
    }
  });

  document.querySelectorAll<HTMLElement>('[data-open]').forEach((el) => {
    el.addEventListener('click', () => {
      const p = el.getAttribute('data-open');
      if (p) void loadFileByPath(p);
    });
  });

  document.querySelectorAll<HTMLElement>('[data-dir]').forEach((el) => {
    el.addEventListener('click', () => {
      const d = el.getAttribute('data-dir');
      if (d) void browseDir(d);
    });
  });

  document.querySelectorAll<HTMLElement>('[data-browse]').forEach((el) => {
    el.addEventListener('click', () => {
      const d = el.getAttribute('data-browse');
      state.drawerOpen = true;
      void browseDir(d || undefined);
    });
  });

  document.querySelectorAll<HTMLElement>('[data-sheet]').forEach((el) => {
    el.addEventListener('click', () => {
      state.activeSheet = Number(el.getAttribute('data-sheet') || 0);
      state.selectedCell = null;
      render();
    });
  });

  document.querySelectorAll<HTMLTableCellElement>('.excel-table td[data-r]').forEach((td) => {
    td.addEventListener('click', () => {
      const r = Number(td.getAttribute('data-r'));
      const c = Number(td.getAttribute('data-c'));
      const sheet = state.doc?.sheets?.[state.activeSheet];
      const val = sheet?.rows?.[r]?.[c] ?? '';
      state.selectedCell = { ref: `${colLetter(c)}${r + 1}`, value: val };
      render();
    });
  });

  // Drag and drop anywhere on the page
  const dropTarget = document.getElementById('dropzone') || document.body;
  dropTarget.ondragover = (e) => {
    e.preventDefault();
    dropTarget.classList.add('dragging');
  };
  dropTarget.ondragleave = () => {
    dropTarget.classList.remove('dragging');
  };
  dropTarget.ondrop = (e) => {
    e.preventDefault();
    dropTarget.classList.remove('dragging');
    const file = e.dataTransfer?.files?.[0];
    if (file) void handleFileUpload(file);
  };
}

if (app) {
  // Register handlers BEFORE app.connect() so no initial notification is missed
  app.ontoolinput = ({ arguments: args }) => {
    const filePath = (args as Record<string, unknown> | undefined)?.filePath;
    if (typeof filePath === 'string' && filePath.trim()) {
      state.loading = true;
      render();
    }
  };

  app.ontoolresult = (result) => {
    const sc = (result.structuredContent || {}) as Record<string, unknown>;
    if (Array.isArray(sc.recent)) state.recent = sc.recent as RecentItem[];
    if (Array.isArray(sc.shortcuts)) state.shortcuts = sc.shortcuts as Shortcut[];

    if (result.isError || sc.error) {
      state.loading = false;
      state.error = String(sc.error || 'Error opening document');
      render();
      return;
    }

    if (sc.mode === 'file' && typeof sc.filePath === 'string') {
      void loadFileByPath(sc.filePath);
    } else {
      state.loading = false;
      render();
    }
  };

  app.onhostcontextchanged = (ctx) => {
    if (ctx.theme) applyDocumentTheme(ctx.theme);
    if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  };
}

async function initSidecar() {
  const params = new URLSearchParams(window.location.search);
  const requestedFile = params.get('file');
  try {
    const res = await callBackend('app_get_recent', {});
    const sc = res.structuredContent || {};
    if (Array.isArray(sc.recent)) state.recent = sc.recent as RecentItem[];
    if (Array.isArray(sc.shortcuts)) state.shortcuts = sc.shortcuts as Shortcut[];
    const active = sc.active as { filePath?: string; updatedAt?: number } | null;
    if (active?.updatedAt) lastActiveUpdatedAt = active.updatedAt;

    if (requestedFile) {
      await loadFileByPath(requestedFile);
    } else if (active?.filePath) {
      await loadFileByPath(active.filePath);
    } else {
      render();
    }
  } catch {
    render();
  }

  // Live-sync when agent opens a new document via open_office_viewer
  window.setInterval(async () => {
    try {
      const res = await callBackend('app_get_recent', {});
      const sc = res.structuredContent || {};
      const active = sc.active as { filePath?: string; updatedAt?: number } | null;
      if (active?.filePath && active.updatedAt && active.updatedAt > lastActiveUpdatedAt) {
        lastActiveUpdatedAt = active.updatedAt;
        if (active.filePath !== state.doc?.filePath) {
          await loadFileByPath(active.filePath);
        }
      }
    } catch {
      // Ignore transient poll errors
    }
  }, 2000);
}

async function init() {
  render();
  if (IS_SIDECAR) {
    await initSidecar();
    return;
  }
  await app!.connect();
  const ctx = app!.getHostContext();
  if (ctx?.theme) applyDocumentTheme(ctx.theme);
  if (ctx?.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  const caps = app!.getHostCapabilities();
  state.canUpdateContext = Boolean(caps?.updateModelContext);
  render();
}

void init();
