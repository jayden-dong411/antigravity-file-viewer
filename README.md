# Antigravity Office Viewer

Native **MCP App** plugin for [Google Antigravity](https://antigravity.google/) that previews local Office documents directly inside the Antigravity side pane or full-page sidebar view — works on **macOS**, **Linux**, and **Windows** without DLL injection or patching the host app.

## Features

- **Word** (`.docx`, `.dotx`): Clean document paper layout with typography, tables, and embedded images.
- **Excel & Spreadsheets** (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`, `.ods`): Multi-sheet tabs, sticky row/column headers, merged cell support, cell value inspector bar, and live row filtering.
- **PowerPoint** (`.pptx`, `.pptm`): Vertical continuous slide cards with titles, bullet hierarchy, tables, embedded images, and toggleable speaker notes.
- **Markdown** (`.md`, `.markdown`): Rendered document preview.
- **Built-in File Navigation**: Path bar, collapsible local folder browser, recent files history, and drag-and-drop / file picker upload.
- **Native Theme Integration**: Automatically adapts to Antigravity light and dark themes via MCP App host CSS variables.

## Installation

### Prerequisites
- **Node.js 20+**
- **Google Antigravity** (with MCP Apps support)

### Quick Install

Clone directly into your Antigravity global plugins directory:

```bash
git clone https://github.com/jayden-dong411/antigravity-office-viewer.git ~/.gemini/config/plugins/office-viewer
cd ~/.gemini/config/plugins/office-viewer
npm install
npm run build
```

Antigravity automatically discovers plugins in `~/.gemini/config/plugins/`.

## Usage

1. **Auto-Preview & One-Click Chat Pills**: Included `rules/AGENTS.md` instructs Antigravity to automatically push generated or inspected local Office files into Office Viewer and display a one-click `sidecar://office-viewer/viewer/?file=...` pill link in chat.
2. **Side Pane (`aux_pane`)**: Click the **`+`** button in the Antigravity side pane and choose **Office Viewer**.
3. **Full Page (`global`)**: Click **Office Viewer** in the left navigation sidebar.
4. **From Chat**: Ask the agent to preview any local file (e.g. `"Open ~/Documents/report.xlsx in Office Viewer"`).

## Development

```bash
npm install
npm run build   # Bundles src/app.ts -> single-file dist/app.html
```

### Project Structure

```text
├── plugin.json              # Antigravity plugin manifest
├── mcp_config.json          # MCP server definition
├── bin/start.sh             # Portable Node.js launcher
├── src/
│   ├── server.mjs           # MCP server (tools + ui:// resource)
│   ├── parsers.mjs          # Office document parsers (.docx/.xlsx/.pptx/.md)
│   └── app.ts               # Interactive MCP App frontend
├── scripts/build.mjs        # Single-file HTML bundler (esbuild)
├── dist/app.html            # Compiled self-contained MCP App HTML
├── skills/office-viewer/    # Agent skill instructions
└── assets/logo.svg          # Plugin icon
```

## License

MIT
