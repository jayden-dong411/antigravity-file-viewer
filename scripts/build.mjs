import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

/* coss ui + Tabler Icons + Impeccable visual tokens & surfaces */
const CSS = `
:root {
  --fv-bg: var(--color-background-primary, #ffffff);
  --fv-surface: var(--color-background-secondary, #f8f9fb);
  --fv-surface-hover: var(--color-background-tertiary, #eff1f5);
  --fv-text: var(--color-text-primary, #111318);
  --fv-text-muted: var(--color-text-secondary, #64748b);
  --fv-border: var(--color-border-primary, #e4e7ec);
  --fv-ring: var(--color-ring-primary, #2563eb);
  --fv-accent: #2563eb;
  --fv-accent-hover: #1d4ed8;
  --fv-radius-sm: 6px;
  --fv-radius-md: 8px;
  --fv-radius-lg: 12px;
  --fv-shadow-xs: 0 1px 2px rgba(16, 24, 40, 0.05);
  --fv-shadow-sm: 0 1px 3px rgba(16, 24, 40, 0.08), 0 1px 2px rgba(16, 24, 40, 0.04);
  --fv-shadow-md: 0 6px 20px rgba(16, 24, 40, 0.1);
  --fv-ease: cubic-bezier(0.2, 0, 0, 1);
}

* { box-sizing: border-box; }
html, body {
  margin: 0;
  padding: 0;
  height: 100%;
  background: var(--fv-bg);
  color: var(--fv-text);
  font-family: var(--font-sans, -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, sans-serif);
  font-size: 13px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}
#app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}

/* Custom thin scrollbars */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb {
  background: color-mix(in srgb, var(--fv-text-muted) 28%, transparent);
  border-radius: 999px;
}
::-webkit-scrollbar-thumb:hover {
  background: color-mix(in srgb, var(--fv-text-muted) 45%, transparent);
}

/* Tabler stroke icons */
.ico {
  width: 14px;
  height: 14px;
  fill: none;
  stroke: currentColor;
  stroke-width: 1.75;
  stroke-linecap: round;
  stroke-linejoin: round;
  flex-shrink: 0;
}

/* Top Toolbar (coss ui layered header) */
.toolbar {
  position: relative;
  border-bottom: 1px solid var(--fv-border);
  background: var(--fv-surface);
  padding: 8px 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  flex-shrink: 0;
  z-index: 10;
}
.toolbar-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}
.toolbar-row.sub {
  justify-content: space-between;
  padding-top: 6px;
  border-top: 1px solid color-mix(in srgb, var(--fv-border) 65%, transparent);
}

/* coss ui Input Group */
.input-group {
  display: flex;
  align-items: center;
  flex: 1;
  min-width: 190px;
  height: 30px;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
  background: var(--fv-bg);
  box-shadow: var(--fv-shadow-xs);
  overflow: hidden;
  transition: border-color 150ms var(--fv-ease), box-shadow 150ms var(--fv-ease);
}
.input-group:focus-within {
  border-color: var(--fv-ring);
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--fv-ring) 16%, transparent);
}
.input-addon {
  display: flex;
  align-items: center;
  padding: 0 6px 0 9px;
  color: var(--fv-text-muted);
}
.input-group input {
  flex: 1;
  min-width: 80px;
  height: 100%;
  border: none;
  background: transparent;
  color: var(--fv-text);
  font-size: 12px;
  padding: 0 6px 0 2px;
  outline: none;
}
.input-submit {
  height: 100%;
  padding: 0 11px;
  border: none;
  border-left: 1px solid var(--fv-border);
  background: var(--fv-surface);
  color: var(--fv-text);
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: background 150ms var(--fv-ease), color 150ms var(--fv-ease);
}
.input-submit:hover {
  background: var(--fv-accent);
  color: #ffffff;
}

/* Search input with leading icon */
.search-wrap {
  display: flex;
  align-items: center;
  gap: 5px;
  height: 28px;
  padding: 0 8px;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
  background: var(--fv-bg);
  color: var(--fv-text-muted);
  transition: border-color 150ms var(--fv-ease), box-shadow 150ms var(--fv-ease);
}
.search-wrap:focus-within {
  border-color: var(--fv-ring);
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--fv-ring) 14%, transparent);
}
.search-input {
  width: 118px;
  border: none;
  background: transparent;
  color: var(--fv-text);
  font-size: 12px;
  outline: none;
}

/* coss ui Buttons */
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 5px;
  height: 30px;
  padding: 0 10px;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
  background: var(--fv-bg);
  color: var(--fv-text);
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  white-space: nowrap;
  box-shadow: var(--fv-shadow-xs);
  transition: background 150ms var(--fv-ease), border-color 150ms var(--fv-ease), transform 100ms var(--fv-ease);
}
.btn:hover { background: var(--fv-surface-hover); }
.btn:active { transform: scale(0.98); }
.btn:focus-visible, .seg-btn:focus-visible, .chip:focus-visible {
  outline: 2px solid var(--fv-ring);
  outline-offset: 1px;
}
.btn.sm { height: 28px; padding: 0 9px; }
.btn.icon-only { width: 28px; padding: 0; }
.btn.ghost {
  border-color: transparent;
  background: transparent;
  box-shadow: none;
}
.btn.ghost:hover { background: var(--fv-surface-hover); }
.btn.primary {
  background: var(--fv-accent);
  border-color: var(--fv-accent);
  color: #ffffff;
  box-shadow: 0 1px 2px rgba(37, 99, 235, 0.28);
}
.btn.primary:hover { background: var(--fv-accent-hover); border-color: var(--fv-accent-hover); }
.btn.active {
  background: color-mix(in srgb, var(--fv-accent) 12%, var(--fv-bg));
  border-color: color-mix(in srgb, var(--fv-accent) 45%, var(--fv-border));
  color: var(--fv-accent);
}

/* coss ui Segmented Control Group */
.seg-group {
  display: inline-flex;
  align-items: center;
  height: 28px;
  padding: 2px;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
  background: color-mix(in srgb, var(--fv-surface-hover) 65%, var(--fv-surface));
  gap: 1px;
}
.seg-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 22px;
  padding: 0 8px;
  border: none;
  border-radius: 5px;
  background: transparent;
  color: var(--fv-text-muted);
  font-size: 11.5px;
  font-weight: 500;
  cursor: pointer;
  white-space: nowrap;
  transition: background 140ms var(--fv-ease), color 140ms var(--fv-ease), box-shadow 140ms var(--fv-ease);
}
.seg-btn.icon-only { width: 24px; padding: 0; }
.seg-btn:hover { color: var(--fv-text); }
.seg-btn.active {
  background: var(--fv-bg);
  color: var(--fv-text);
  box-shadow: var(--fv-shadow-xs);
}
.zoom-label {
  font-size: 11px;
  font-variant-numeric: tabular-nums;
  min-width: 38px;
  text-align: center;
  color: var(--fv-text-muted);
  user-select: none;
}

.doc-meta, .controls {
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
  flex-wrap: wrap;
}
.doc-title {
  font-size: 12.5px;
  font-weight: 600;
  letter-spacing: -0.01em;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 260px;
}
.meta-dot {
  width: 3px;
  height: 3px;
  border-radius: 50%;
  background: var(--fv-text-muted);
  opacity: 0.5;
}
.muted { color: var(--fv-text-muted); font-size: 11.5px; }

/* Subtle top loading progress bar */
.loading-bar {
  position: absolute;
  left: 0;
  right: 0;
  bottom: -1px;
  height: 2px;
  overflow: hidden;
  background: color-mix(in srgb, var(--fv-accent) 18%, transparent);
}
.loading-bar-inner {
  width: 40%;
  height: 100%;
  background: var(--fv-accent);
  animation: fv-slide 1.1s infinite ease-in-out;
}
@keyframes fv-slide {
  0% { transform: translateX(-100%); }
  100% { transform: translateX(350%); }
}

.banner {
  padding: 8px 14px;
  font-size: 12px;
  border-bottom: 1px solid var(--fv-border);
}
.banner.error {
  background: color-mix(in srgb, #dc2626 10%, var(--fv-bg));
  color: #dc2626;
}

.workspace {
  display: flex;
  flex: 1;
  min-height: 0;
  overflow: hidden;
}

/* Sidebar Drawer */
.drawer {
  width: 268px;
  max-width: 45%;
  border-right: 1px solid var(--fv-border);
  background: var(--fv-surface);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
}
.drawer-head {
  padding: 8px 10px;
  border-bottom: 1px solid var(--fv-border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
}
.drawer-path {
  font-size: 11px;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  color: var(--fv-text-muted);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  flex: 1;
}
.drawer-actions { display: flex; align-items: center; gap: 4px; flex-shrink: 0; }
.drawer-list {
  flex: 1;
  overflow-y: auto;
  padding: 6px;
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.drawer-item {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  text-align: left;
  padding: 6px 8px;
  border: none;
  border-radius: var(--fv-radius-sm);
  background: transparent;
  color: var(--fv-text);
  cursor: pointer;
  font-size: 12px;
  transition: background 120ms var(--fv-ease);
}
.drawer-item:hover { background: var(--fv-surface-hover); }
.drawer-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.icon-dir { display: inline-flex; color: var(--fv-text-muted); }
.drawer-chev { display: inline-flex; color: var(--fv-text-muted); opacity: 0.55; }

.viewport {
  flex: 1;
  min-width: 0;
  overflow: auto;
  display: flex;
  flex-direction: column;
}

/* Soft semantic tinted badges (coss ui style) */
.badge {
  display: inline-flex;
  align-items: center;
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 0.03em;
  padding: 2px 6px;
  border-radius: 5px;
  flex-shrink: 0;
  border: 1px solid transparent;
}
.badge-word  { background: rgba(37, 99, 235, 0.11); color: #2563eb; border-color: rgba(37, 99, 235, 0.22); }
.badge-excel { background: rgba(22, 163, 74, 0.11); color: #15803d; border-color: rgba(22, 163, 74, 0.22); }
.badge-ppt   { background: rgba(234, 88, 12, 0.11); color: #c2410c; border-color: rgba(234, 88, 12, 0.22); }
.badge-pdf   { background: rgba(220, 38, 38, 0.11); color: #b91c1c; border-color: rgba(220, 38, 38, 0.22); }
.badge-html  { background: rgba(2, 132, 199, 0.11); color: #0369a1; border-color: rgba(2, 132, 199, 0.22); }
.badge-zip   { background: rgba(124, 58, 237, 0.11); color: #6d28d9; border-color: rgba(124, 58, 237, 0.22); }
.badge-ipynb { background: rgba(249, 115, 22, 0.11); color: #c2410c; border-color: rgba(249, 115, 22, 0.22); }
.badge-media { background: rgba(219, 39, 119, 0.11); color: #be185d; border-color: rgba(219, 39, 119, 0.22); }
.badge-img   { background: rgba(13, 148, 136, 0.11); color: #0f766e; border-color: rgba(13, 148, 136, 0.22); }
.badge-md    { background: rgba(99, 102, 241, 0.11); color: #4338ca; border-color: rgba(99, 102, 241, 0.22); }

/* Welcome / Empty State (coss ui Frame + Card) */
.welcome {
  max-width: 660px;
  width: 100%;
  margin: 28px auto;
  padding: 0 18px 32px;
  display: flex;
  flex-direction: column;
  gap: 22px;
}
.frame {
  padding: 6px;
  border-radius: 16px;
  background: var(--fv-surface);
  border: 1px solid var(--fv-border);
}
.dropzone {
  border: 1.5px dashed color-mix(in srgb, var(--fv-border) 90%, var(--fv-text-muted));
  border-radius: var(--fv-radius-lg);
  padding: 32px 24px;
  text-align: center;
  background: var(--fv-bg);
  box-shadow: var(--fv-shadow-xs);
  transition: border-color 160ms var(--fv-ease), background 160ms var(--fv-ease);
}
.dropzone.dragging {
  border-color: var(--fv-accent);
  background: color-mix(in srgb, var(--fv-accent) 5%, var(--fv-bg));
}
.hero-icon {
  width: 44px;
  height: 44px;
  margin: 0 auto 12px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: color-mix(in srgb, var(--fv-accent) 10%, var(--fv-surface));
  border: 1px solid color-mix(in srgb, var(--fv-accent) 22%, var(--fv-border));
  color: var(--fv-accent);
}
.hero-icon .ico { width: 20px; height: 20px; }
.drop-title {
  font-size: 18px;
  font-weight: 650;
  letter-spacing: -0.02em;
  margin-bottom: 6px;
}
.drop-sub {
  color: var(--fv-text-muted);
  max-width: 460px;
  margin: 0 auto 14px;
  font-size: 12.5px;
}
.format-pills {
  display: flex;
  justify-content: center;
  flex-wrap: wrap;
  gap: 5px;
  margin-bottom: 18px;
}
.drop-actions { display: flex; justify-content: center; gap: 8px; margin-bottom: 16px; }
.shortcuts { display: flex; justify-content: center; flex-wrap: wrap; gap: 6px; }
.chip {
  border: 1px solid var(--fv-border);
  background: var(--fv-surface);
  color: var(--fv-text-muted);
  border-radius: 999px;
  padding: 4px 11px;
  font-size: 11.5px;
  font-weight: 500;
  cursor: pointer;
  transition: all 140ms var(--fv-ease);
}
.chip:hover {
  color: var(--fv-text);
  border-color: color-mix(in srgb, var(--fv-accent) 45%, var(--fv-border));
  background: var(--fv-bg);
}
.section-title {
  font-weight: 600;
  margin-bottom: 8px;
  color: var(--fv-text-muted);
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.06em;
}
.file-list { display: flex; flex-direction: column; gap: 6px; }
.file-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 9px 12px;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
  background: var(--fv-bg);
  color: var(--fv-text);
  cursor: pointer;
  text-align: left;
  box-shadow: var(--fv-shadow-xs);
  transition: border-color 140ms var(--fv-ease), background 140ms var(--fv-ease), transform 120ms var(--fv-ease);
}
.file-item:hover {
  border-color: color-mix(in srgb, var(--fv-accent) 45%, var(--fv-border));
  background: var(--fv-surface);
}
.file-info { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 1px; }
.file-name { font-weight: 600; font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.file-path { font-size: 11px; color: var(--fv-text-muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.file-size { font-size: 11px; color: var(--fv-text-muted); font-variant-numeric: tabular-nums; flex-shrink: 0; }

/* Word & Markdown paper */
.doc-canvas {
  padding: 24px 18px 48px;
  background: var(--fv-surface);
  min-height: 100%;
}
.paper {
  max-width: 820px;
  margin: 0 auto;
  padding: 36px 42px;
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  border-radius: var(--fv-radius-lg);
  box-shadow: var(--fv-shadow-sm);
  overflow-wrap: break-word;
  line-height: 1.7;
}
.paper h1, .paper h2, .paper h3 { letter-spacing: -0.02em; line-height: 1.3; }
.paper img { max-width: 100%; height: auto; border-radius: 6px; }
.paper table { border-collapse: collapse; width: 100%; margin: 14px 0; font-size: 0.94em; }
.paper th, .paper td { border: 1px solid var(--fv-border); padding: 7px 10px; text-align: left; }
.paper th { background: var(--fv-surface); font-weight: 600; }
.paper pre { background: var(--fv-surface); border: 1px solid var(--fv-border); padding: 12px; border-radius: var(--fv-radius-md); overflow-x: auto; }
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
  padding: 6px 10px;
  background: var(--fv-surface);
  border-bottom: 1px solid var(--fv-border);
  overflow-x: auto;
  flex-shrink: 0;
}
.sheet-tab {
  border: 1px solid transparent;
  background: transparent;
  color: var(--fv-text-muted);
  padding: 4px 11px;
  border-radius: var(--fv-radius-sm);
  cursor: pointer;
  font-size: 12px;
  font-weight: 500;
  white-space: nowrap;
  transition: all 120ms var(--fv-ease);
}
.sheet-tab:hover { color: var(--fv-text); background: var(--fv-surface-hover); }
.sheet-tab.active {
  background: var(--fv-bg);
  color: var(--fv-text);
  border-color: var(--fv-border);
  box-shadow: var(--fv-shadow-xs);
  font-weight: 600;
}
.formula-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 5px 12px;
  background: var(--fv-surface);
  border-bottom: 1px solid var(--fv-border);
  font-size: 12px;
  flex-shrink: 0;
}
.cell-ref {
  font-weight: 700;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  color: #16a34a;
  min-width: 38px;
}
.cell-val { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sheet-wrap { flex: 1; overflow: auto; }
.excel-table { border-collapse: collapse; width: max-content; min-width: 100%; }
.excel-table th, .excel-table td {
  border: 1px solid var(--fv-border);
  padding: 5px 10px;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.excel-table thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--fv-surface);
  color: var(--fv-text-muted);
  font-size: 11.5px;
  font-weight: 600;
  text-align: center;
}
.excel-table .row-num {
  position: sticky;
  left: 0;
  z-index: 1;
  background: var(--fv-surface);
  color: var(--fv-text-muted);
  font-variant-numeric: tabular-nums;
  font-size: 11px;
  text-align: right;
  min-width: 38px;
  user-select: none;
}
.excel-table thead .row-num { z-index: 3; }
.excel-table tbody tr:hover td:not(.row-num) {
  background: color-mix(in srgb, var(--fv-accent) 4%, transparent);
}
.excel-table tr.first-row td:not(.row-num) { font-weight: 600; }
.excel-table td.cell-match { background: rgba(250, 204, 21, 0.32); }
.sheet-meta {
  padding: 5px 12px;
  font-size: 11px;
  color: var(--fv-text-muted);
  border-top: 1px solid var(--fv-border);
  background: var(--fv-surface);
  flex-shrink: 0;
}

/* PPTX canvas (pptx-preview) */
.pptx-canvas {
  padding: 24px 18px 44px;
  background: var(--fv-surface);
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
  margin: 0 auto 20px auto !important;
  border-radius: var(--fv-radius-md);
  box-shadow: var(--fv-shadow-md);
  border: 1px solid color-mix(in srgb, var(--fv-border) 80%, transparent);
  overflow: hidden;
  background: #ffffff;
}
.pptx-notes-panel {
  width: 100%;
  max-width: 960px;
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  border-radius: var(--fv-radius-md);
  padding: 14px 18px;
  box-shadow: var(--fv-shadow-xs);
}
.pptx-note-item {
  font-size: 12px;
  margin-top: 6px;
  color: var(--fv-text-muted);
}

/* PDF & HTML iframes */
.pdf-stage {
  flex: 1;
  display: flex;
  min-height: 0;
  height: 100%;
  background: var(--fv-surface);
}
.pdf-frame { width: 100%; height: 100%; border: none; }
.html-stage {
  flex: 1;
  display: flex;
  justify-content: center;
  min-height: 0;
  height: 100%;
  background: var(--fv-surface);
  overflow: auto;
}
.html-frame {
  height: 100%;
  border: none;
  background: #ffffff;
  box-shadow: var(--fv-shadow-md);
}
.html-source-wrap { flex: 1; overflow: auto; padding: 14px; background: var(--fv-surface); }
.code-view {
  margin: 0;
  padding: 16px;
  border-radius: var(--fv-radius-md);
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 12px;
  line-height: 1.6;
  white-space: pre-wrap;
  word-break: break-word;
  overflow: auto;
}

/* ZIP archive explorer */
.zip-layout {
  display: flex;
  flex: 1;
  min-height: 0;
  height: 100%;
  overflow: hidden;
}
.zip-pane {
  width: 340px;
  max-width: 48%;
  border-right: 1px solid var(--fv-border);
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: var(--fv-surface);
}
.zip-summary {
  padding: 8px 12px;
  font-size: 11px;
  color: var(--fv-text-muted);
  border-bottom: 1px solid var(--fv-border);
  flex-shrink: 0;
}
.zip-list {
  flex: 1;
  overflow-y: auto;
  padding: 6px;
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.zip-item {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  padding: 6px 8px;
  border: none;
  border-radius: var(--fv-radius-sm);
  background: transparent;
  color: var(--fv-text);
  font-size: 12px;
  text-align: left;
  transition: background 120ms var(--fv-ease);
}
.zip-item.file { cursor: pointer; }
.zip-item.file:hover { background: var(--fv-surface-hover); }
.zip-item.file.active {
  background: color-mix(in srgb, var(--fv-accent) 12%, var(--fv-bg));
  color: var(--fv-accent);
  font-weight: 600;
}
.zip-item.dir { color: var(--fv-text-muted); }
.zip-path { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.zip-preview {
  flex: 1;
  min-width: 0;
  overflow: auto;
  padding: 14px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  background: var(--fv-bg);
}
.zip-preview-head {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--fv-border);
}
.zip-img-wrap { display: flex; justify-content: center; padding: 12px; }
.zip-img-wrap img {
  max-width: 100%;
  max-height: 75vh;
  border-radius: var(--fv-radius-md);
  border: 1px solid var(--fv-border);
}

/* Jupyter Notebook (.ipynb) */
.nb-canvas {
  padding: 20px 18px 44px;
  background: var(--fv-surface);
  min-height: 100%;
}
.nb-header {
  max-width: 900px;
  margin: 0 auto 14px;
  padding: 10px 14px;
  border-radius: var(--fv-radius-md);
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  box-shadow: var(--fv-shadow-xs);
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.nb-cells {
  max-width: 900px;
  margin: 0 auto;
  display: flex;
  flex-direction: column;
  gap: 12px;
}
.nb-cell {
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  border-radius: var(--fv-radius-md);
  box-shadow: var(--fv-shadow-xs);
  overflow: hidden;
}
.nb-md, .nb-raw, .nb-code-row { display: flex; align-items: stretch; }
.nb-gutter {
  width: 68px;
  flex-shrink: 0;
  padding: 10px 8px;
  text-align: right;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 11px;
  color: var(--fv-text-muted);
  background: var(--fv-surface);
  border-right: 1px solid var(--fv-border);
  user-select: none;
}
.nb-md-body { flex: 1; min-width: 0; padding: 10px 16px; }
.nb-md-body img { max-width: 100%; }
.nb-code {
  flex: 1;
  min-width: 0;
  margin: 0;
  padding: 10px 14px;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 12px;
  line-height: 1.55;
  overflow-x: auto;
  white-space: pre;
}
.nb-outputs {
  border-top: 1px solid var(--fv-border);
  padding: 10px 14px 10px 82px;
  background: var(--fv-bg);
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.nb-out { margin: 0; font-size: 12px; overflow-x: auto; }
.nb-out-stream, .nb-out-text { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; white-space: pre-wrap; }
.nb-out-stream.stderr { color: #b45309; }
.nb-out-err {
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  white-space: pre-wrap;
  padding: 8px 10px;
  border-radius: var(--fv-radius-sm);
  background: rgba(220, 38, 38, 0.08);
  color: #dc2626;
}
.nb-out-img img { max-width: 100%; border-radius: 6px; }
.nb-out-html table { border-collapse: collapse; font-size: 12px; }
.nb-out-html th, .nb-out-html td { border: 1px solid var(--fv-border); padding: 4px 8px; }

/* Media (Video / Audio) */
.media-stage {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  min-height: 0;
  height: 100%;
  background: #0c0e13;
}
.video-player {
  max-width: 100%;
  max-height: 100%;
  border-radius: var(--fv-radius-md);
  box-shadow: 0 12px 36px rgba(0, 0, 0, 0.5);
  background: #000000;
}
.audio-card {
  width: 100%;
  max-width: 520px;
  padding: 26px;
  border-radius: var(--fv-radius-lg);
  background: var(--fv-bg);
  border: 1px solid var(--fv-border);
  box-shadow: var(--fv-shadow-md);
  display: flex;
  flex-direction: column;
  gap: 10px;
  text-align: center;
}
.audio-title { font-size: 16px; font-weight: 650; word-break: break-all; }
.audio-player { width: 100%; margin-top: 8px; }

/* Enhanced Image & SVG viewer */
.img-stage {
  flex: 1;
  min-height: 0;
  overflow: auto;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
}
.img-stage.bg-checker {
  background-color: #e5e7eb;
  background-image:
    linear-gradient(45deg, #cbd5e1 25%, transparent 25%),
    linear-gradient(-45deg, #cbd5e1 25%, transparent 25%),
    linear-gradient(45deg, transparent 75%, #cbd5e1 75%),
    linear-gradient(-45deg, transparent 75%, #cbd5e1 75%);
  background-size: 20px 20px;
  background-position: 0 0, 0 10px, 10px -10px, -10px 0;
}
.img-stage.bg-light { background: #ffffff; }
.img-stage.bg-dark { background: #111318; }
.preview-img {
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
  transition: transform 140ms var(--fv-ease);
  transform-origin: center center;
}
.empty-inline { padding: 28px; text-align: center; color: var(--fv-text-muted); }

@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    transition-duration: 0.01ms !important;
  }
}
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
    --color-background-primary: #111318;
    --color-background-secondary: #181b22;
    --color-background-tertiary: #232732;
    --color-text-primary: #eef2f6;
    --color-text-secondary: #94a3b8;
    --color-border-primary: #2a2f3b;
  }
  .badge-word  { color: #60a5fa; }
  .badge-excel { color: #4ade80; }
  .badge-ppt   { color: #fb923c; }
  .badge-pdf   { color: #f87171; }
  .badge-html  { color: #38bdf8; }
  .badge-zip   { color: #a78bfa; }
  .badge-ipynb { color: #fb923c; }
  .badge-media { color: #f472b6; }
  .badge-img   { color: #2dd4bf; }
  .badge-md    { color: #818cf8; }
}
`;

const makeHtml = (mode) => `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>File Viewer</title>
  <style>${mode === 'sidecar' ? darkModeFallback : ''}${CSS}</style>
</head>
<body>
  <div id="app"></div>
  ${mode === 'sidecar' ? '<script>window.__FILE_VIEWER_MODE__="sidecar";</script>' : ''}
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
