# Outlook Web MCP

Search and read Outlook Web from Codex or Claude Code. The local MCP server uses Playwright and returns compact results. It operates through the visible Outlook UI and requires an account you are authorized to use with your agent.

```sh
npm ci
npx playwright install chromium
npm run login -- work
npm run login -- work2
```

Sign in and complete MFA in each Chromium window. Each window closes when the mailbox is ready. Use `personal` for an Outlook.com or Hotmail account. Other aliases open Microsoft 365 Outlook. Each alias has its own Chromium profile under `~/.outlook-web-mcp/profiles/`. Treat those directories as credentials.

Register the MCP using the absolute path to the repository:

```sh
codex mcp add outlook-web -- node /absolute/path/to/outlook-web-mcp/server.mjs
claude mcp add --scope user outlook-web -- node /absolute/path/to/outlook-web-mcp/server.mjs
```

Start a new agent session. `search_mail` takes `account` (defaults to `work`), `query`, and `limit` (defaults to 5). It returns sender, subject, date, and an ID. `read_mail` takes an ID and `max_chars` (defaults to 4,000). It uses the account associated with that ID. Opening unread mail can mark it as read.

The server starts Chromium Headless Shell on demand and closes each browser after one minute of inactivity. Search IDs expire when that browser closes or you search that account again. Use a given alias in one agent session at a time. Run the login command again if Microsoft expires the session.

The current selectors target the English Outlook Web UI. Outlook UI changes can require selector updates. Email text returned by the MCP enters your agent's context.
