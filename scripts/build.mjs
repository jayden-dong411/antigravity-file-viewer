import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const CSS = `
* { box-sizing: border-box; }
html, body {
  margin: 0;
  padding: 0;
  height: 100%;
  background: var(--color-background-primary, #ffffff);
  color: var(--color-text-primary, #1f2328);
  font-family: var(--font-sans, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif);
  font-size: 13px;
  line-height: 1.5;
}
#app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}
.toolbar {
  border-bottom: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-secondary, #f6f8fa);
  padding: 8px 10px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  flex-shrink: 0;
}
.toolbar-row {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
}
.toolbar-row.sub {
  justify-content: space-between;
}
.path-form {
  display: flex;
  flex: 1;
  min-width: 180px;
  gap: 6px;
}
input[type="text"], input[type="search"] {
  flex: 1;
  min-width: 100px;
  height: 28px;
  padding: 0 8px;
  border-radius: 6px;
  border: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-primary, #ffffff);
  color: var(--color-text-primary, #1f2328);
  font-size: 12px;
  outline: none;
}
input[type="text"]:focus, input[type="search"]:focus {
  border-color: var(--color-ring-primary, #2563eb);
}
.search-input {
  width: 140px;
  flex: 0 1 140px;
}
.btn {
  height: 28px;
  padding: 0 10px;
  border-radius: 6px;
  border: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-primary, #ffffff);
  color: var(--color-text-primary, #1f2328);
  font-size: 12px;
  cursor: pointer;
  white-space: nowrap;
}
.btn:hover { background: var(--color-background-tertiary, #eaeef2); }
.btn.sm { height: 26px; padding: 0 8px; }
.btn.primary {
  background: #2563eb;
  border-color: #2563eb;
  color: #ffffff;
}
.btn.primary:hover { background: #1d4ed8; }
.btn.active {
  background: var(--color-background-tertiary, #dbeafe);
  border-color: #2563eb;
}
.doc-meta, .controls {
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
}
.doc-title {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 240px;
}
.muted { color: var(--color-text-secondary, #656d76); font-size: 11px; }
.zoom-label { font-size: 11px; min-width: 34px; text-align: center; color: var(--color-text-secondary, #656d76); }
.banner {
  padding: 6px 12px;
  font-size: 12px;
  border-bottom: 1px solid var(--color-border-primary, #d0d7de);
}
.banner.error { background: var(--color-background-danger, #ffebe9); color: var(--color-text-danger, #cf222e); }
.banner.info { background: var(--color-background-info, #ddf4ff); color: var(--color-text-info, #0969da); }
.workspace {
  display: flex;
  flex: 1;
  min-height: 0;
  overflow: hidden;
}
.drawer {
  width: 260px;
  max-width: 45%;
  border-right: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-secondary, #f6f8fa);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
}
.drawer-head {
  padding: 8px;
  border-bottom: 1px solid var(--color-border-primary, #d0d7de);
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.drawer-path {
  font-size: 11px;
  color: var(--color-text-secondary, #656d76);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.drawer-actions { display: flex; gap: 6px; }
.drawer-list {
  flex: 1;
  overflow-y: auto;
  padding: 4px;
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.drawer-item {
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  text-align: left;
  padding: 6px 8px;
  border: none;
  border-radius: 6px;
  background: transparent;
  color: var(--color-text-primary, #1f2328);
  cursor: pointer;
  font-size: 12px;
}
.drawer-item:hover { background: var(--color-background-tertiary, #eaeef2); }
.drawer-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.icon-dir {
  font-size: 10px;
  font-weight: 700;
  padding: 1px 5px;
  border-radius: 4px;
  background: var(--color-background-tertiary, #e5e7eb);
  color: var(--color-text-secondary, #4b5563);
}
.viewport {
  flex: 1;
  min-width: 0;
  overflow: auto;
  display: flex;
  flex-direction: column;
}
.badge {
  display: inline-block;
  font-size: 10px;
  font-weight: 700;
  padding: 1px 5px;
  border-radius: 4px;
  color: #fff;
  flex-shrink: 0;
}
.badge-word { background: #2563eb; }
.badge-excel { background: #16a34a; }
.badge-ppt { background: #ea580c; }
.badge-md { background: #6366f1; }

.welcome {
  max-width: 680px;
  width: 100%;
  margin: 24px auto;
  padding: 0 16px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}
.dropzone {
  border: 2px dashed var(--color-border-primary, #d0d7de);
  border-radius: 12px;
  padding: 28px 20px;
  text-align: center;
  background: var(--color-background-secondary, #f6f8fa);
  transition: border-color 0.15s;
}
.dropzone.dragging { border-color: #2563eb; }
.drop-title { font-size: 18px; font-weight: 700; margin-bottom: 6px; }
.drop-sub { color: var(--color-text-secondary, #656d76); margin-bottom: 16px; }
.drop-actions { display: flex; justify-content: center; gap: 8px; margin-bottom: 14px; }
.shortcuts { display: flex; justify-content: center; flex-wrap: wrap; gap: 6px; }
.chip {
  border: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-primary, #fff);
  color: var(--color-text-secondary, #656d76);
  border-radius: 999px;
  padding: 3px 10px;
  font-size: 11px;
  cursor: pointer;
}
.chip:hover { color: var(--color-text-primary, #1f2328); border-color: #2563eb; }
.section-title { font-weight: 600; margin-bottom: 8px; color: var(--color-text-secondary, #656d76); font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; }
.file-list { display: flex; flex-direction: column; gap: 6px; }
.file-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 10px;
  border-radius: 8px;
  border: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-secondary, #f6f8fa);
  color: var(--color-text-primary, #1f2328);
  cursor: pointer;
  text-align: left;
}
.file-item:hover { border-color: #2563eb; }
.file-info { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.file-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.file-path { font-size: 11px; color: var(--color-text-secondary, #656d76); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.file-size { font-size: 11px; color: var(--color-text-secondary, #656d76); flex-shrink: 0; }

/* Word & Markdown paper */
.doc-canvas {
  padding: 20px 16px 40px;
  background: var(--color-background-secondary, #f3f4f6);
  min-height: 100%;
}
.paper {
  max-width: 820px;
  margin: 0 auto;
  padding: 32px 36px;
  background: var(--color-background-primary, #ffffff);
  border: 1px solid var(--color-border-primary, #d0d7de);
  border-radius: 8px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  overflow-wrap: break-word;
}
.paper img { max-width: 100%; height: auto; border-radius: 4px; }
.paper table { border-collapse: collapse; width: 100%; margin: 12px 0; }
.paper th, .paper td { border: 1px solid var(--color-border-primary, #d0d7de); padding: 6px 8px; text-align: left; }
.paper pre { background: var(--color-background-secondary, #f6f8fa); padding: 10px; border-radius: 6px; overflow-x: auto; }
.paper code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 0.9em; }

/* Spreadsheet */
.sheet-container {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
}
.sheet-tabs {
  display: flex;
  gap: 4px;
  padding: 6px 8px;
  background: var(--color-background-secondary, #f6f8fa);
  border-bottom: 1px solid var(--color-border-primary, #d0d7de);
  overflow-x: auto;
  flex-shrink: 0;
}
.sheet-tab {
  border: 1px solid transparent;
  background: transparent;
  color: var(--color-text-secondary, #656d76);
  padding: 4px 10px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 12px;
  white-space: nowrap;
}
.sheet-tab.active {
  background: var(--color-background-primary, #ffffff);
  color: var(--color-text-primary, #1f2328);
  border-color: var(--color-border-primary, #d0d7de);
  font-weight: 600;
}
.formula-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 10px;
  background: var(--color-background-secondary, #f6f8fa);
  border-bottom: 1px solid var(--color-border-primary, #d0d7de);
  font-size: 12px;
  flex-shrink: 0;
}
.cell-ref {
  font-weight: 700;
  color: #16a34a;
  min-width: 36px;
}
.cell-val { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sheet-wrap { flex: 1; overflow: auto; }
.excel-table { border-collapse: collapse; width: max-content; min-width: 100%; }
.excel-table th, .excel-table td {
  border: 1px solid var(--color-border-primary, #d0d7de);
  padding: 4px 8px;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.excel-table thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--color-background-secondary, #f6f8fa);
  font-weight: 600;
  text-align: center;
}
.excel-table .row-num {
  position: sticky;
  left: 0;
  z-index: 1;
  background: var(--color-background-secondary, #f6f8fa);
  color: var(--color-text-secondary, #656d76);
  text-align: right;
  min-width: 36px;
  user-select: none;
}
.excel-table thead .row-num { z-index: 3; }
.excel-table tr.first-row td:not(.row-num) { font-weight: 600; }
.excel-table td.cell-match { background: rgba(250, 204, 21, 0.35); }
.sheet-meta {
  padding: 4px 10px;
  font-size: 11px;
  color: var(--color-text-secondary, #656d76);
  border-top: 1px solid var(--color-border-primary, #d0d7de);
  background: var(--color-background-secondary, #f6f8fa);
  flex-shrink: 0;
}

/* PPTX canvas (pptx-preview) */
.pptx-canvas {
  padding: 20px 16px 40px;
  background: var(--color-background-secondary, #eef1f5);
  min-height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 18px;
}
.pptx-stage {
  display: flex;
  flex-direction: column;
  align-items: center;
  width: 100%;
}
.pptx-stage .pptx-preview-wrapper {
  background: transparent !important;
}
.pptx-stage .pptx-preview-slide-wrapper {
  margin: 0 auto 18px auto !important;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
  overflow: hidden;
  background: #ffffff;
}
.pptx-notes-panel {
  width: 100%;
  max-width: 960px;
  background: var(--color-background-primary, #ffffff);
  border: 1px solid var(--color-border-primary, #d0d7de);
  border-radius: 8px;
  padding: 14px 18px;
}
.pptx-note-item {
  font-size: 12px;
  margin-top: 6px;
  color: var(--color-text-secondary, #656d76);
}
.empty-inline { padding: 24px; text-align: center; color: var(--color-text-secondary, #656d76); }
`;

const result = await esbuild.build({
  entryPoints: [path.join(ROOT, 'src', 'app.ts')],
  bundle: true,
  minify: true,
  format: 'iife',
  target: ['es2022'],
  write: false,
});

const jsCode = result.outputFiles[0].text.replace(/<\/script>/gi, '<\\/script>');
const darkModeFallback = `
@media (prefers-color-scheme: dark) {
  :root {
    --color-background-primary: #16181d;
    --color-background-secondary: #1e2128;
    --color-background-tertiary: #282c34;
    --color-text-primary: #e6edf3;
    --color-text-secondary: #9aa4b2;
    --color-border-primary: #30363d;
  }
}
`;

const makeHtml = (mode) => `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Office Viewer</title>
  <style>${mode === 'sidecar' ? darkModeFallback : ''}${CSS}</style>
</head>
<body>
  <div id="app"></div>
  ${mode === 'sidecar' ? '<script>window.__OFFICE_VIEWER_MODE__="sidecar";</script>' : ''}
  <script>${jsCode}</script>
</body>
</html>`;

await fs.mkdir(path.join(ROOT, 'dist'), { recursive: true });
const appHtml = makeHtml('mcp-app');
const sidecarHtml = makeHtml('sidecar');
await fs.writeFile(path.join(ROOT, 'dist', 'app.html'), appHtml, 'utf8');
await fs.writeFile(path.join(ROOT, 'dist', 'sidecar.html'), sidecarHtml, 'utf8');
console.log(
  `Built dist/app.html (${(Buffer.byteLength(appHtml) / 1024).toFixed(1)} KB) and dist/sidecar.html (${(
    Buffer.byteLength(sidecarHtml) / 1024
  ).toFixed(1)} KB)`,
);
