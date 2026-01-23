# Teams MCP Server

An MCP (Model Context Protocol) server that enables AI assistants to search Microsoft Teams messages via browser automation and direct API calls.

## Why Browser Automation?

Microsoft Graph API requires complex authentication setup, Azure AD app registration, and specific permissions. This project sidesteps that by automating the Teams web app directly, using your existing browser session.

## How It Works

1. **First search**: Opens browser → you log in → search runs → auth tokens captured
2. **Subsequent searches**: Uses cached tokens to call the Substrate API directly (no browser)
3. **Token expiry**: Automatically falls back to browser when tokens expire (~1 hour)

## Prerequisites

- Node.js 18+
- A Microsoft account with Teams access

## Installation

```bash
npm install
npx playwright install chromium
```

## Usage

### CLI Tools

```bash
# Check authentication status
npm run cli -- status

# Search (opens browser for login if needed)
npm run cli -- search "meeting notes"

# Search with pagination
npm run cli -- search "project" --from 0 --size 50

# Force browser mode (skip direct API)
npm run cli -- search "query" --browser

# Output as JSON
npm run cli -- search "query" --json
```

### MCP Server

Add to your MCP client configuration:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": ["/path/to/team-mcp/dist/index.js"]
    }
  }
}
```

Or for development:

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["tsx", "/path/to/team-mcp/src/index.ts"]
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `teams_search` | Search messages with query, pagination (from, size, maxResults) |
| `teams_login` | Trigger manual login (visible browser) |
| `teams_status` | Check authentication and session state |

## Session Management

- **Session state**: `./session-state.json` (gitignored)
- **Token cache**: `./token-cache.json` (gitignored)
- **Browser profile**: `./.user-data/` (gitignored)

If your session expires, call `teams_login` or delete these files and search again.

## Development

```bash
# Run MCP server in development mode
npm run dev

# Build for production
npm run build

# Research/explore Teams APIs
npm run research
```

## Limitations

- Requires initial manual login
- May break if Microsoft changes the Teams web app
- Tokens expire after ~1 hour (auto-fallback to browser)

---

## Teams Chat Export Bookmarklet

This repo also includes a standalone bookmarklet for exporting Teams chat messages to Markdown.

See [teams-bookmarklet/README.md](teams-bookmarklet/README.md) for details.
