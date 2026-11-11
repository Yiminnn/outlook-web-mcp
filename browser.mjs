import { chromium } from 'playwright';
import { mkdir } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

export const profileUrl = account => account === 'personal'
  ? 'https://outlook.live.com/mail/0/'
  : 'https://outlook.cloud.microsoft/mail/';
export const searchBox = page => page.getByRole('combobox', { name: /Search for email/ });
const outlookHosts = ['outlook.cloud.microsoft', 'outlook.office.com', 'outlook.office365.com', 'outlook.live.com'];

function isOutlookUrl(value) {
  const url = new URL(value);
  return url.protocol === 'https:' && outlookHosts.some(host =>
    url.hostname === host || url.hostname === host + '.mcas.ms');
}

export const isOutlook = page => isOutlookUrl(page.url());

export async function openProfile(account, headless) {
  if (!/^[a-z][a-z0-9_-]{0,31}$/.test(account)) throw new Error('Invalid profile alias.');
  const root = join(homedir(), '.outlook-web-mcp', 'profiles');
  const directory = join(root, account);
  await mkdir(root, { recursive: true, mode: 0o700 });
  await mkdir(directory, { recursive: true, mode: 0o700 });
  const context = await chromium.launchPersistentContext(directory, {
    headless, chromiumSandbox: true, timeout: 15000,
  });
  const page = context.pages()[0] ?? await context.newPage();
  return { context, page };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const account = process.argv[2];
  if (!account) throw new Error('Usage: npm run login -- <alias>');
  const { context, page } = await openProfile(account, false);
  try {
    await page.goto(profileUrl(account), {waitUntil: 'domcontentloaded'});
    console.log(`Sign in to Outlook in the ${account} window. It closes when the mailbox is ready.`);
    const deadline = Date.now() + 300000;
    while (!page.isClosed() && Date.now() < deadline) {
      if (isOutlook(page) && await searchBox(page).isVisible() && !new URL(page.url()).hash) break;
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    if (page.isClosed() || !isOutlook(page) || !await searchBox(page).isVisible()) {
      throw new Error('Outlook mailbox not reached before the sign-in window expired.');
    }
    console.log(`Saved profile ${account}.`);
  } finally {
    await context.close();
  }
}
