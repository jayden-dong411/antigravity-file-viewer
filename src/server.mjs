import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import {
  browseDirectory,
  expandPath,
  loadOfficeFile,
  parseOfficeBuffer,
} from './parsers.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN_ROOT = path.resolve(__dirname, '..');
const APP_URI = 'ui://office-viewer/app.html';
const APP_HTML_PATH = process.env.OFFICE_VIEWER_APP_HTML || path.join(PLUGIN_ROOT, 'dist', 'app.html');

function parseDataDir() {
  const arg = process.argv.find((a) => a.startsWith('--data-dir='));
  if (arg) return arg.slice('--data-dir='.length);
  return path.join(os.homedir(), '.gemini', 'antigravity', 'plugin_data', 'office-viewer');
}

const DATA_DIR = parseDataDir();
const RECENT_PATH = path.join(DATA_DIR, 'recent.json');
const ACTIVE_PATH = path.join(DATA_DIR, 'active.json');

async function readRecent() {
  try {
    const raw = await fs.readFile(RECENT_PATH, 'utf8');
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.slice(0, 15) : [];
  } catch {
    return [];
  }
}

async function addRecent(entry) {
  try {
    await fs.mkdir(DATA_DIR, { recursive: true });
    const list = await readRecent();
    const filtered = list.filter((item) => item.filePath !== entry.filePath);
    const next = [
      {
        filePath: entry.filePath,
        fileName: entry.fileName,
        ext: entry.ext,
        size: entry.size,
        openedAt: Date.now(),
      },
      ...filtered,
    ].slice(0, 15);
    await fs.writeFile(RECENT_PATH, JSON.stringify(next, null, 2), 'utf8');
    await fs.writeFile(
      ACTIVE_PATH,
      JSON.stringify({ filePath: entry.filePath, fileName: entry.fileName, updatedAt: Date.now() }, null, 2),
      'utf8',
    );
    return next;
  } catch {
    return [];
  }
}

const server = new Server(
  {
    name: 'office-viewer',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
      resources: {},
    },
  },
);

server.setRequestHandler(ListResourcesRequestSchema, async () => ({
  resources: [
    {
      uri: APP_URI,
      name: 'Office Viewer App',
      description: 'Interactive Office document viewer for Word, Excel, PowerPoint, and Markdown files',
      mimeType: 'text/html;profile=mcp-app',
      _meta: {
        ui: {
          prefersBorder: false,
        },
      },
    },
  ],
}));

server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  if (request.params.uri !== APP_URI) {
    throw new Error(`Unknown resource: ${request.params.uri}`);
  }
  const html = await fs.readFile(APP_HTML_PATH, 'utf8');
  return {
    contents: [
      {
        uri: APP_URI,
        mimeType: 'text/html;profile=mcp-app',
        text: html,
      },
    ],
  };
});

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'open_office_viewer',
      title: 'Office Viewer',
      description:
        'Open the interactive Office Viewer in the Antigravity side pane to preview local Word (.docx), Excel (.xlsx/.xls/.csv/.tsv/.ods), PowerPoint (.pptx), or Markdown (.md) files. Call with no arguments to open the file browser and recent files view.',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: {
            type: 'string',
            description: 'Optional absolute or ~-prefixed path to a local Office document to preview immediately.',
          },
        },
      },
      _meta: {
        ui: {
          resourceUri: APP_URI,
        },
        'com.google.antigravity/ui': {
          entrypoints: [
            { type: 'aux_pane', title: 'Office Viewer' },
            { type: 'global', title: 'Office Viewer' },
          ],
        },
      },
    },
    {
      name: 'app_load_file',
      description: 'App-only tool: parse and return rendered data for a local Office document.',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: { type: 'string' },
          maxRows: { type: 'number' },
        },
        required: ['filePath'],
      },
      _meta: {
        ui: {
          visibility: ['app'],
        },
      },
    },
    {
      name: 'app_parse_upload',
      description: 'App-only tool: parse a drag-and-dropped or uploaded Office document sent as base64.',
      inputSchema: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          base64: { type: 'string' },
          maxRows: { type: 'number' },
        },
        required: ['fileName', 'base64'],
      },
      _meta: {
        ui: {
          visibility: ['app'],
        },
      },
    },
    {
      name: 'app_browse_directory',
      description: 'App-only tool: list folders and supported Office files inside a local directory.',
      inputSchema: {
        type: 'object',
        properties: {
          dirPath: { type: 'string' },
        },
      },
      _meta: {
        ui: {
          visibility: ['app'],
        },
      },
    },
    {
      name: 'app_get_recent',
      description: 'App-only tool: return recently opened Office documents and quick folder shortcuts.',
      inputSchema: {
        type: 'object',
        properties: {},
      },
      _meta: {
        ui: {
          visibility: ['app'],
        },
      },
    },
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  try {
    if (name === 'open_office_viewer') {
      const recent = await readRecent();
      const homeDir = os.homedir();
      const shortcuts = [
        { label: 'Home', path: homeDir },
        { label: 'Desktop', path: path.join(homeDir, 'Desktop') },
        { label: 'Documents', path: path.join(homeDir, 'Documents') },
        { label: 'Downloads', path: path.join(homeDir, 'Downloads') },
      ];

      const rawPath = typeof args.filePath === 'string' ? args.filePath.trim() : '';
      if (!rawPath) {
        return {
          content: [{ type: 'text', text: 'Opened Office Viewer. Select or drop a local .docx, .xlsx, .csv, .pptx, or .md file to preview.' }],
          structuredContent: {
            mode: 'home',
            recent,
            shortcuts,
          },
        };
      }

      const resolved = expandPath(rawPath);
      const stat = await fs.stat(resolved);
      if (!stat.isFile()) {
        throw new Error(`Not a file: ${resolved}`);
      }
      const fileName = path.basename(resolved);
      const ext = path.extname(fileName).toLowerCase();
      const updatedRecent = await addRecent({ filePath: resolved, fileName, ext, size: stat.size });

      return {
        content: [{ type: 'text', text: `Opened "${fileName}" (${resolved}) in Office Viewer.` }],
        structuredContent: {
          mode: 'file',
          filePath: resolved,
          fileName,
          ext,
          size: stat.size,
          recent: updatedRecent,
          shortcuts,
        },
      };
    }

    if (name === 'app_load_file') {
      const data = await loadOfficeFile(args.filePath, { maxRows: args.maxRows });
      const recent = await addRecent(data);
      return {
        content: [{ type: 'text', text: `Loaded ${data.fileName}` }],
        structuredContent: { ...data, recent },
      };
    }

    if (name === 'app_parse_upload') {
      const buffer = Buffer.from(String(args.base64 || ''), 'base64');
      const fileName = String(args.fileName || 'document');
      const parsed = await parseOfficeBuffer(buffer, fileName, { maxRows: args.maxRows });
      return {
        content: [{ type: 'text', text: `Parsed uploaded file ${fileName}` }],
        structuredContent: {
          filePath: '',
          fileName,
          ext: path.extname(fileName).toLowerCase(),
          size: buffer.length,
          ...parsed,
        },
      };
    }

    if (name === 'app_browse_directory') {
      const listing = await browseDirectory(args.dirPath);
      return {
        content: [{ type: 'text', text: `Listed ${listing.dirPath}` }],
        structuredContent: listing,
      };
    }

    if (name === 'app_get_recent') {
      const recent = await readRecent();
      const homeDir = os.homedir();
      return {
        content: [{ type: 'text', text: `Returned ${recent.length} recent files` }],
        structuredContent: {
          recent,
          shortcuts: [
            { label: 'Home', path: homeDir },
            { label: 'Desktop', path: path.join(homeDir, 'Desktop') },
            { label: 'Documents', path: path.join(homeDir, 'Documents') },
            { label: 'Downloads', path: path.join(homeDir, 'Downloads') },
          ],
        },
      };
    }

    throw new Error(`Unknown tool: ${name}`);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      isError: true,
      content: [{ type: 'text', text: `Office Viewer error: ${message}` }],
      structuredContent: { error: message },
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
