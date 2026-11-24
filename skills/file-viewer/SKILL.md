---
name: file-viewer
description: Preview local Office documents (.docx, .xlsx, .xls, .csv, .tsv, .ods, .pptx), PDF (.pdf), HTML (.html/.htm), ZIP archives (.zip), and Markdown (.md) directly inside the Antigravity side pane using the File Viewer MCP App and Sidecar.
---

# File Viewer

Use the `file-viewer` MCP server's `open_file_viewer` tool whenever the user wants to preview or inspect a local Word (`.docx`), Excel/spreadsheet (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.ods`), PowerPoint (`.pptx`), PDF (`.pdf`), HTML (`.html`, `.htm`), ZIP archive (`.zip`), or Markdown (`.md`) file directly inside Antigravity.

## Usage

- Call `open_file_viewer` with `{ "filePath": "/absolute/path/to/file.docx" }` to open the document directly in the Antigravity side pane.
- Call `open_file_viewer` with `{}` (empty arguments) to open the File Viewer browser, recent files list, and drag-and-drop uploader in the side pane.
- Note: Respect any user rule that prefers Google Slides for `.ppt`/`.pptx` unless the user explicitly asks to preview it locally in File Viewer.
