# Outlook MCP

Search and read email from Codex or Claude Code. Choose the Outlook for Mac server to use the signed-in desktop app, or the web server to use Outlook in Playwright. Use accounts you are authorized to access with your agent.

## Outlook for Mac

### Install

Requires macOS, Outlook for Mac, Node.js 18 or later, and the Xcode Command Line Tools. Open Outlook, sign in to both accounts, and expand each account in the mail sidebar.

```sh
git clone https://github.com/Yiminnn/outlook-web-mcp.git
cd outlook-web-mcp
npm ci
swiftc desktop.swift -o .outlook-desktop-ax
```

In the same directory, run the command for each client you use.

**Codex**

```sh
codex mcp add outlook-desktop -- "$(command -v node)" "$PWD/desktop.mjs"
codex mcp get outlook-desktop
```

**Claude Code**

```sh
claude mcp add --scope user outlook-desktop -- "$(command -v node)" "$PWD/desktop.mjs"
claude mcp get outlook-desktop
```

Start a new client session. If the helper reports that Accessibility is required, allow the app that runs the MCP under **System Settings → Privacy & Security → Accessibility**. Outlook must be open. Use one client at a time because both clients operate the same Outlook window.

### Search and read two accounts

1. Call `search_mail` with the exact address Outlook shows for the first account in its sidebar. Replace the example address with your own. `limit` defaults to 5.
   ```json
   {"account":"first@example.com","query":"seminar","limit":5}
   ```
2. Call `read_mail` with an ID from those results. `max_chars` defaults to 4,000 and accepts up to 50,000.
   ```json
   {"id":"00000000-0000-4000-8000-000000000001","max_chars":4000}
   ```
3. To switch accounts, call `search_mail` again with the second account's sidebar address, then use an ID from the new results.
   ```json
   {"account":"second@example.com","query":"newsletter","limit":5}
   ```

The helper checks the window title and `Current Mailbox` search scope. Each search invalidates earlier result IDs in that client session. The server brings Outlook to the foreground and reads plain text exposed by the accessibility tree. The `truncated` flag indicates whether the character limit was reached. Opening an unread message may mark it as read. The current implementation expects English Outlook labels. Email text returned by the MCP enters your agent's context.

## Outlook Web

The web server uses Playwright and keeps each browser profile separate.

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

The server starts Chromium Headless Shell on demand and closes each browser after one minute of inactivity. Search IDs expire when that browser closes or you search that account again. Use a given alias in one agent session at a time. Run the login command again if Microsoft expires the session. Some organizations require fresh authentication when the browser restarts. In that case this setup cannot resume those sessions headlessly.

The current selectors target the English Outlook Web UI. Outlook UI changes can require selector updates. Email text returned by the MCP enters your agent's context.
