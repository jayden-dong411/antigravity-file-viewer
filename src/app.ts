import {
  App,
  applyDocumentTheme,
  applyHostStyleVariables,
} from '@modelcontextprotocol/ext-apps';

type SheetMerge = { sr: number; sc: number; er: number; ec: number };
type SheetData = {
  name: string;
  rows: string[][];
  merges: SheetMerge[];
  totalRows: number;
  totalCols: number;
  truncated: boolean;
};

type SlideBlock =
  | { type: 'text' | 'subtitle'; paragraphs: { text: string; level: number }[] }
  | { type: 'table'; rows: string[][] };

type SlideData = {
  index: number;
  title: string;
  blocks: SlideBlock[];
  images: string[];
  notes: string;
};

type RecentItem = {
  filePath: string;
  fileName: string;
  ext: string;
  size: number;
  openedAt: number;
};

type Shortcut = { label: string; path: string };

type LoadedDoc = {
  filePath: string;
  fileName: string;
  ext: string;
  size: number;
  kind: 'docx' | 'sheet' | 'pptx' | 'markdown';
  html?: string;
  warnings?: string[];
  sheets?: SheetData[];
  slides?: SlideData[];
  recent?: RecentItem[];
};

const app = new App({ name: 'Office Viewer', version: '1.0.0' }, {});

const state: {
  doc: LoadedDoc | null;
  activeSheet: number;
  zoom: number;
  search: string;
  showNotes: boolean;
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
    const res = await app.callServerTool({
      name: 'app_load_file',
      arguments: { filePath: filePath.trim() },
    });
    const sc = (res.structuredContent || {}) as Record<string, unknown>;
    if (res.isError || sc.error) {
      state.error = String(sc.error || 'Failed to load file');
    } else {
      state.doc = sc as unknown as LoadedDoc;
      state.activeSheet = 0;
      state.search = '';
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
    const res = await app.callServerTool({
      name: 'app_parse_upload',
      arguments: { fileName: file.name, base64 },
    });
    const sc = (res.structuredContent || {}) as Record<string, unknown>;
    if (res.isError || sc.error) {
      state.error = String(sc.error || 'Failed to parse file');
    } else {
      state.doc = sc as unknown as LoadedDoc;
      state.activeSheet = 0;
      state.search = '';
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
    const res = await app.callServerTool({
      name: 'app_browse_directory',
      arguments: dirPath ? { dirPath } : {},
    });
    const sc = (res.structuredContent || {}) as Record<string, unknown>;
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

function renderPptx(slides: SlideData[]): string {
  if (!slides || slides.length === 0) {
    return `<div class="empty-inline">No slides found in this presentation.</div>`;
  }
  const query = state.search.trim().toLowerCase();
  const filtered = query
    ? slides.filter(
        (s) =>
          s.title.toLowerCase().includes(query) ||
          s.notes.toLowerCase().includes(query) ||
          s.blocks.some((b) =>
            b.type === 'table'
              ? b.rows.some((r) => r.some((c) => c.toLowerCase().includes(query)))
              : b.paragraphs.some((p) => p.text.toLowerCase().includes(query)),
          ),
      )
    : slides;

  const cards = filtered
    .map((slide) => {
      const blocksHtml = slide.blocks
        .map((b) => {
          if (b.type === 'table') {
            const rowsHtml = b.rows
              .map(
                (r, idx) =>
                  `<tr>${r
                    .map((c) => (idx === 0 ? `<th>${esc(c)}</th>` : `<td>${esc(c)}</td>`))
                    .join('')}</tr>`,
              )
              .join('');
            return `<table class="slide-table">${rowsHtml}</table>`;
          }
          const cls = b.type === 'subtitle' ? 'slide-subtitle' : 'slide-list';
          const items = b.paragraphs
            .map((p) => `<div class="slide-p lvl-${Math.min(p.level, 3)}">${esc(p.text)}</div>`)
            .join('');
          return `<div class="${cls}">${items}</div>`;
        })
        .join('');

      const imgsHtml = slide.images.length
        ? `<div class="slide-images">${slide.images.map((src) => `<img src="${src}" alt="Slide media" />`).join('')}</div>`
        : '';

      const notesHtml =
        state.showNotes && slide.notes
          ? `<div class="slide-notes"><strong>Speaker Notes:</strong><div>${esc(slide.notes).replace(/\n/g, '<br>')}</div></div>`
          : '';

      return `
        <article class="slide-card">
          <div class="slide-head">
            <span class="slide-num">Slide ${slide.index}</span>
            <h2 class="slide-title">${esc(slide.title)}</h2>
          </div>
          <div class="slide-body">${blocksHtml}${imgsHtml}</div>
          ${notesHtml}
        </article>`;
    })
    .join('');

  return `<div class="slides-stack" style="zoom:${state.zoom}%">${cards}</div>`;
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
          <div class="drop-title">Office Viewer</div>
          <div class="drop-sub">Drop a local <code>.docx</code>, <code>.xlsx</code>, <code>.csv</code>, <code>.pptx</code>, or <code>.md</code> file here, click to choose, or browse folders.</div>
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
    return renderPptx(doc.slides || []);
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
        <input id="file-input" type="file" accept=".docx,.dotx,.xlsx,.xls,.xlsm,.xlsb,.csv,.tsv,.ods,.pptx,.pptm,.md,.markdown" hidden />
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
                <input id="search-input" class="search-input" type="search" value="${esc(state.search)}" placeholder="Filter / search..." />
                ${doc.kind === 'pptx' ? `<button class="btn sm ${state.showNotes ? 'active' : ''}" id="notes-btn">Notes</button>` : ''}
                <button class="btn sm" id="zoom-out">-</button>
                <span class="zoom-label">${state.zoom}%</span>
                <button class="btn sm" id="zoom-in">+</button>
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

  document.getElementById('attach-btn')?.addEventListener('click', async () => {
    if (!state.doc?.filePath) return;
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

async function init() {
  render();
  await app.connect();
  const ctx = app.getHostContext();
  if (ctx?.theme) applyDocumentTheme(ctx.theme);
  if (ctx?.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  const caps = app.getHostCapabilities();
  state.canUpdateContext = Boolean(caps?.updateModelContext);
  render();
}

void init();
