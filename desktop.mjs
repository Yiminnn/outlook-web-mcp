import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';

const helper = fileURLToPath(new URL('./.outlook-desktop-ax', import.meta.url));
const server = new McpServer({ name: 'outlook-desktop', version: '1.0.0' });
const results = new Map();
let pending = Promise.resolve();

function invoke(input) {
  return new Promise(resolve => {
    const child = spawn(helper, [], { stdio: ['pipe', 'pipe', 'ignore'] });
    let output = '';
    const timer = setTimeout(() => child.kill(), 60000);
    child.stdout.setEncoding('utf8');
    child.stdout.on('data', chunk => { output += chunk; });
    child.stdin.on('error', () => {});
    child.once('error', () => resolve({ error: 'Build the local Outlook helper before running this MCP.' }));
    child.once('close', () => {
      clearTimeout(timer);
      try { resolve(JSON.parse(output)); }
      catch { resolve({ error: 'The Outlook helper timed out or returned an invalid response.' }); }
    });
    child.stdin.end(JSON.stringify(input));
  });
}

function respond(operation) {
  const task = pending.then(operation).catch(() => ({ error: 'The Outlook operation failed.' }));
  pending = task;
  return task.then(result => ({
    ...(result.error ? { isError: true } : {}),
    content: [{ type: 'text', text: JSON.stringify(result) }],
  }));
}

server.registerTool('search_mail', {
  description: 'Search only the specified account in Outlook for Mac.',
  inputSchema: {
    account: z.string().email().max(254),
    query: z.string().trim().min(1).max(400),
    limit: z.number().int().min(1).max(10).default(5),
  },
}, input => respond(async () => {
  results.clear();
  const result = await invoke({ operation: 'search', ...input });
  if (result.error) return result;
  return { matches: result.rows.map((row, index) => {
    const id = randomUUID();
    results.set(id, { account: input.account, query: input.query, row, index });
    const [metadata, remainder] = row.trim().split(/,\s{4,}/, 2);
    return { id, metadata: remainder === undefined ? null : metadata.slice(0, 350) };
  }) };
}));

server.registerTool('read_mail', {
  description: 'Read a result from the latest scoped Outlook search. Opening unread mail may mark it as read.',
  inputSchema: {
    id: z.string().uuid(),
    max_chars: z.number().int().min(1).max(50000).default(4000),
  },
}, ({ id, max_chars }) => respond(() => {
  const result = results.get(id);
  return result ? invoke({ operation: 'read', ...result, max_chars })
    : { error: 'Use an ID from the most recent search in this client session.' };
}));

await server.connect(new StdioServerTransport());
