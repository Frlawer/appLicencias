import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(moduleDir, '..');
const contextFile = path.join(moduleDir, 'project-context.md');

const ignoredDirectories = new Set(['.git', 'node_modules', '.clasp', '.mcp', 'dist', 'build', 'out']);
const ignoredFiles = new Set(['package-lock.json']);
const textExtensions = new Set(['.gs', '.js', '.html', '.json', '.md', '.txt', '.css']);

const server = new McpServer({
  name: 'app-licencias',
  version: '1.0.0'
});

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function readText(filePath) {
  return fs.readFile(filePath, 'utf8');
}

async function collectFiles(currentDir, relativeBase = '') {
  const entries = await fs.readdir(currentDir, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    if (ignoredDirectories.has(entry.name)) continue;
    const absolutePath = path.join(currentDir, entry.name);
    const relativePath = path.join(relativeBase, entry.name);

    if (entry.isDirectory()) {
      files.push(...await collectFiles(absolutePath, relativePath));
      continue;
    }

    if (ignoredFiles.has(entry.name)) continue;
    files.push(relativePath.replace(/\\/g, '/'));
  }

  return files;
}

async function searchInProject(query) {
  const normalizedQuery = String(query || '').trim().toLowerCase();
  if (!normalizedQuery) return [];

  const files = await collectFiles(projectRoot);
  const matches = [];

  for (const relativePath of files) {
    const absolutePath = path.join(projectRoot, relativePath);
    const extension = path.extname(relativePath).toLowerCase();
    if (!textExtensions.has(extension)) continue;

    let content;
    try {
      content = await readText(absolutePath);
    } catch {
      continue;
    }

    const lowerContent = content.toLowerCase();
    const index = lowerContent.indexOf(normalizedQuery);
    if (index === -1) continue;

    const before = content.slice(0, index);
    const line = before.split(/\r?\n/).length;
    matches.push({ path: relativePath, line });
    if (matches.length >= 25) break;
  }

  return matches;
}

server.registerResource('project-summary', 'applicencias://summary', {
  mimeType: 'text/markdown',
  description: 'Resumen y reglas del proyecto App Licencias'
}, async () => ({
  contents: [{
    uri: 'applicencias://summary',
    text: await readText(contextFile)
  }]
}));

server.registerResource('project-files', 'applicencias://files', {
  mimeType: 'application/json',
  description: 'Listado de archivos del proyecto App Licencias'
}, async () => ({
  contents: [{
    uri: 'applicencias://files',
    text: JSON.stringify(await collectFiles(projectRoot), null, 2)
  }]
}));

server.registerTool('get_project_context', {
  description: 'Devuelve el contexto curado del proyecto y el listado completo de archivos',
  inputSchema: {
    type: 'object',
    properties: {}
  }
}, async () => {
  const summary = await readText(contextFile);
  const files = await collectFiles(projectRoot);

  return {
    content: [{
      type: 'text',
      text: [
        '# Contexto del proyecto',
        '',
        summary,
        '',
        '## Archivos detectados',
        ...files.map(file => `- ${file}`)
      ].join('\n')
    }]
  };
});

server.registerTool('list_project_files', {
  description: 'Lista archivos del proyecto, opcionalmente filtrados por prefijo',
  inputSchema: {
    type: 'object',
    properties: {
      pathPrefix: { type: 'string' }
    }
  }
}, async ({ pathPrefix }) => {
  const files = await collectFiles(projectRoot);
  const prefix = String(pathPrefix || '').replace(/\\/g, '/');
  const filtered = prefix ? files.filter(file => file.startsWith(prefix)) : files;

  return {
    content: [{
      type: 'text',
      text: JSON.stringify(filtered, null, 2)
    }]
  };
});

server.registerTool('read_project_file', {
  description: 'Lee un archivo del proyecto usando una ruta relativa segura',
  inputSchema: {
    type: 'object',
    properties: {
      path: { type: 'string' }
    },
    required: ['path']
  }
}, async ({ path: requestedPath }) => {
  const normalizedRelativePath = String(requestedPath || '').replace(/^([/\\])+/, '').replace(/\\/g, '/');
  const absolutePath = path.resolve(projectRoot, normalizedRelativePath);
  const relativeToRoot = path.relative(projectRoot, absolutePath);

  if (relativeToRoot.startsWith('..') || path.isAbsolute(relativeToRoot)) {
    return {
      content: [{ type: 'text', text: 'Ruta inválida.' }],
      isError: true
    };
  }

  if (!(await exists(absolutePath))) {
    return {
      content: [{ type: 'text', text: `No se encontró el archivo: ${normalizedRelativePath}` }],
      isError: true
    };
  }

  const content = await readText(absolutePath);
  return {
    content: [{
      type: 'text',
      text: content
    }]
  };
});

server.registerTool('search_project_text', {
  description: 'Busca texto dentro de los archivos de texto del proyecto',
  inputSchema: {
    type: 'object',
    properties: {
      query: { type: 'string' }
    },
    required: ['query']
  }
}, async ({ query }) => {
  const matches = await searchInProject(query);
  return {
    content: [{
      type: 'text',
      text: JSON.stringify(matches, null, 2)
    }]
  };
});

const transport = new StdioServerTransport();
await server.connect(transport);