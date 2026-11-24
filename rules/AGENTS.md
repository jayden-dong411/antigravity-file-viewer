# File Viewer Auto-Preview & Sidecar Pill Rules

Whenever you create, update, locate, or are asked to inspect a local Office, PDF, HTML, or archive document (`.docx`, `.dotx`, `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`, `.ods`, `.pdf`, `.html`, `.htm`, `.zip`, or `.pptx` when the user explicitly requests local preview rather than Google Slides):

1. **Auto-push to File Viewer**: Call the `file-viewer` MCP tool `open_file_viewer` with `{ "filePath": "/absolute/path/to/file" }` so the document opens/updates directly in the Antigravity side pane and syncs as the active preview file.
2. **Surface a one-click sidecar pill link**: In your chat response, include a clickable `sidecar://` pill link so the user can open or focus that exact document in the File Viewer side pane with one click:
   - `[Preview <filename>](sidecar://file-viewer/viewer/?file=<url-encoded-absolute-path>)`
   - Or, when opening the viewer without a specific file: `[File Viewer](sidecar://file-viewer/viewer/)`
3. Respect any user rule that prioritizes Google Slides for `.ppt`/`.pptx` files unless the user explicitly asks to preview locally in File Viewer.
