---
name: office-viewer
description: Preview local Office documents (.docx, .xlsx, .xls, .csv, .tsv, .ods, .pptx, .md) directly inside the Antigravity side pane using the Office Viewer MCP App.
---

# Office Viewer

Use the `office-viewer` MCP server's `open_office_viewer` tool whenever the user wants to preview or inspect a local Word (`.docx`), Excel/spreadsheet (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.ods`), PowerPoint (`.pptx`), or Markdown (`.md`) file directly inside Antigravity.

## Usage

- Call `open_office_viewer` with `{ "filePath": "/absolute/path/to/file.docx" }` to open the document directly in the Antigravity side pane.
- Call `open_office_viewer` with `{}` (empty arguments) to open the Office Viewer browser, recent files list, and drag-and-drop uploader in the side pane.
- Note: Respect any user rule that prefers Google Slides for `.ppt`/`.pptx` unless the user explicitly asks to preview it locally in Office Viewer.
