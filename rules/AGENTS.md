# Office Viewer Auto-Preview & Sidecar Pill Rules

Whenever you create, update, locate, or are asked to inspect a local Office document (`.docx`, `.dotx`, `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`, `.ods`, or `.pptx` when the user explicitly requests local preview rather than Google Slides):

1. **Auto-push to Office Viewer**: Call the `office-viewer` MCP tool `open_office_viewer` with `{ "filePath": "/absolute/path/to/file" }` so the document opens/updates directly in the Antigravity side pane and syncs as the active preview file.
2. **Surface a one-click sidecar pill link**: In your chat response, include a clickable `sidecar://` pill link so the user can open or focus that exact document in the Office Viewer side pane with one click:
   - `[Preview <filename>](sidecar://office-viewer/viewer/?file=<url-encoded-absolute-path>)`
   - Or, when opening the viewer without a specific file: `[Office Viewer](sidecar://office-viewer/viewer/)`
3. Respect any user rule that prioritizes Google Slides for `.ppt`/`.pptx` files unless the user explicitly asks to preview locally in Office Viewer.
