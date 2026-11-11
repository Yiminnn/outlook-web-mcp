import { randomUUID } from 'node:crypto';
import { isOutlook, openProfile, profileUrl, searchBox } from './browser.mjs';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';

const profiles = new Map();
const alias = z.string().regex(/^[a-z][a-z0-9_-]{0,31}$/);
const idleTimeout = 60_000;
const server = new McpServer({ name: 'outlook-web-mail', version: '1.0.0' });
const resultIds = new Map();
let pending = Promise.resolve();

async function browser(account, operation, args) {
  let state = profiles.get(account);
  try {
    if (!state) {
      state = await openProfile(account, true);
      profiles.set(account, state);
      await state.page.goto(profileUrl(account),
        { waitUntil: 'domcontentloaded', timeout: 15000 });
    }
    clearTimeout(state.timer);
    await searchBox(state.page).waitFor({ timeout: 10000 });
    if (!isOutlook(state.page)) throw new Error('Unexpected Outlook origin.');
    return await operation(state.page, args);
  } catch {
    return { error: 'Outlook is unavailable. Check that the profile is signed in, free in other clients, and on the expected Outlook page.' };
  } finally {
    if (state) {
      state.timer = setTimeout(() => {
        profiles.delete(account);
        for (const [id, result] of resultIds) {
          if (result.account === account) resultIds.delete(id);
        }
        state.context.close().catch(() => {});
      }, idleTimeout).unref();
    }
  }
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
  description: 'Search Outlook Web. Returns metadata and temporary IDs for the first visible results.',
  inputSchema: {
    account: alias.default('work'),
    query: z.string().trim().min(1).max(400),
    limit: z.number().int().min(1).max(10).default(5),
  },
}, ({ account, query, limit }) => respond(async () => {
  for (const [id, result] of resultIds) {
    if (result.account === account) resultIds.delete(id);
  }
  const result = await browser(account, async (page, { query, limit }) => {
    const search = page.getByRole('combobox', { name: /Search for email/ });
    await search.click({ timeout: 12000 });
    const sameQuery = await search.inputValue() === query;
    const rows = page.getByRole('listbox', { name: 'Message list' })
      .locator('[role=option][data-convid][data-focusable-row]');
    const empty = page.getByText(/^(No results found|No results|We couldn.t find anything|We didn.t find anything)/i);
    const existing = rows.first().or(empty.first()).first();
    const previous = !sameQuery && await existing.count() ? await existing.elementHandle() : null;
    await search.fill(query, { timeout: 12000 });
    await search.press('Enter');
    await page.getByRole('heading', { name: 'Results', exact: true }).first()
      .waitFor({ timeout: 10000 });
    if (previous) {
      await page.waitForFunction(element => !element.isConnected, previous, { timeout: 10000 });
    }
    await existing.waitFor({ timeout: 10000 });
    const matches = await rows.evaluateAll((items, limit) => items.slice(0, limit).map(row => ({
      key: row.getAttribute('data-convid'),
      from: row.querySelector('span[title]')?.getAttribute('title')?.slice(0, 200) ?? null,
      subject: row.querySelector('.TtcXM')?.innerText.trim().slice(0, 250) ?? null,
      date: row.querySelector('.qq2gS[title]')?.getAttribute('title')?.slice(0, 80) ?? null,
    })), limit);
    return { matches };
  }, { query, limit });

  if (result.error) return result;
  return { matches: result.matches.map(({ key, ...fields }) => {
    const id = randomUUID();
    resultIds.set(id, { account, key });
    return { id, ...fields };
  }) };
}));

server.registerTool('read_mail', {
  description: 'Read visible mail text using an ID from the latest search. Opening unread mail may mark it as read.',
  inputSchema: {
    id: z.string().uuid(),
    max_chars: z.number().int().min(1).max(12000).default(4000),
  },
}, ({ id, max_chars }) => respond(async () => {
  const result = resultIds.get(id);
  if (!result) return { error: 'Use a mail ID from the latest search in this client session.' };
  return browser(result.account, async (page, { key, max_chars }) => {
    const selector = await page.evaluate(key => `[role=option][data-convid="${CSS.escape(key)}"]`, key);
    const row = page.getByRole('listbox', { name: 'Message list' }).locator(selector);
    if (await row.count() !== 1) return { error: 'The search view changed. Search for the message again.' };
    const body = page.locator('[role=document]:visible').last();
    const previous = await body.count() ? await body.elementHandle() : null;
    const alreadySelected = await row.getAttribute('aria-selected') === 'true';
    await row.click({ timeout: 8000 });
    await row.and(page.locator('[aria-selected=true]')).waitFor({ timeout: 8000 });
    if (previous && !alreadySelected) {
      await page.waitForFunction(element => !element.isConnected, previous, { timeout: 8000 });
    }
    await body.waitFor({ timeout: 10000 });
    return body.evaluate((element, limit) => {
      const clone = element.cloneNode(true);
      clone.querySelectorAll('blockquote, #divRplyFwdMsg, [id="appendonsend"]').forEach(node => {
        if (node.matches('#divRplyFwdMsg, [id="appendonsend"]')) {
          while (node.nextSibling) node.nextSibling.remove();
        }
        node.remove();
      });
      const text = clone.isEqualNode(element) ? element.innerText : clone.textContent;
      const clean = text.replace(/\n{3,}/g, '\n\n').trim();
      return { body: clean.slice(0, limit), truncated: clean.length > limit };
    }, max_chars);
  }, { key: result.key, max_chars });
}));

await server.connect(new StdioServerTransport());
