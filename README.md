# Teams MCP Server

An MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams-search messages, send replies, manage favourites, and more.

## How It Works

This server calls Microsoft's internal Teams APIs directly (Substrate, chatsvc, CSA)-the same APIs the Teams web app uses. No Azure AD app registration or admin consent required.

**Authentication flow:**
1. First use opens a browser for you to log in
2. OAuth tokens are extracted and cached
3. Subsequent calls use cached tokens directly (no browser)
4. Tokens auto-refresh when expired (~1 hour); session cookies keep you logged in longer

## Installation

### Prerequisites

- Node.js 18+
- A Microsoft account with Teams access
- Google Chrome, Microsoft Edge, or Chromium browser installed

### Setup

```bash
git clone https://github.com/your-org/team-mcp.git
cd team-mcp
PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1 npm install
npm run build
```

The server uses your system's installed Chrome (macOS/Linux) or Edge (Windows) for authentication. This avoids downloading Playwright's bundled Chromium (~180MB).

**If you don't have Chrome/Edge installed**, you can download Playwright's browser instead:

```bash
npm install
npx playwright install chromium
npm run build
```

### Configure Your MCP Client

Add to your MCP client configuration (e.g., Claude Desktop, Cursor):

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

For development (uses tsx for hot reload):

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

### Search & Discovery

| Tool | Description |
|------|-------------|
| `teams_search` | Search messages with operators (`from:`, `sent:`, `in:`, `hasattachment:`, etc.) |
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_find_channel` | Find channels by name (your teams + org-wide discovery) |

### Messaging

| Tool | Description |
|------|-------------|
| `teams_send_message` | Send a message (default: self-chat/notes) |
| `teams_reply_to_thread` | Reply to a channel message as a threaded reply |
| `teams_edit_message` | Edit one of your own messages |
| `teams_delete_message` | Delete one of your own messages (soft delete) |

### People & Contacts

| Tool | Description |
|------|-------------|
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (useful for name resolution) |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |

### Organisation

| Tool | Description |
|------|-------------|
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark from a message |

### Session

| Tool | Description |
|------|-------------|
| `teams_login` | Trigger manual login (opens browser) |
| `teams_status` | Check authentication and session state |

### Search Operators

The search supports Teams' native operators:

```
from:sarah@company.com     # Messages from person
sent:today                 # Messages from today
sent:lastweek              # Messages from last week
in:project-alpha           # Messages in channel
"Rob Smith"                # Find @mentions (name in quotes)
hasattachment:true         # Messages with files
NOT from:email@co.com      # Exclude results
```

Combine operators: `from:sarah@co.com sent:lastweek hasattachment:true`

**Note:** `@me`, `from:me`, `to:me` do NOT work. Use `teams_get_me` first to get your email/displayName, then use those values.

## MCP Resources

The server also exposes passive resources for context discovery:

| Resource URI | Description |
|--------------|-------------|
| `teams://me/profile` | Current user's profile |
| `teams://me/favorites` | Pinned conversations |
| `teams://status` | Authentication status |

## CLI Tools

A command-line interface is included for testing and debugging:

```bash
# Check authentication status
npm run cli -- status

# Search messages
npm run cli -- search "meeting notes"
npm run cli -- search "project" --from 0 --size 50

# Send messages
npm run cli -- send "Hello from Teams MCP!"
npm run cli -- send "Message" --to "conversation-id"

# Force login
npm run cli -- login --force

# Output as JSON
npm run cli -- search "query" --json
```

### MCP Test Harness

Test the server through the actual MCP protocol:

```bash
# List available tools
npm run test:mcp

# Call any tool
npm run test:mcp -- search "your query"
npm run test:mcp -- status
npm run test:mcp -- people "john smith"
npm run test:mcp -- favorites
```

## Limitations

- **Initial login required** - First use opens a browser for manual authentication
- **Undocumented APIs** - Uses Microsoft's internal APIs which may change without notice
- **Token refresh** - Opens browser briefly every ~1 hour to refresh tokens; manual re-login only needed if session cookies expire
- **Search limitations** - Full-text search only; thread replies not matching search terms won't appear (use `teams_get_thread` for full context)
- **Own messages only** - Edit/delete only works on your own messages

## Session Files

These files are created locally and gitignored:

- `session-state.json` - Encrypted browser session
- `token-cache.json` - Encrypted OAuth tokens
- `.user-data/` - Browser profile

If your session expires, call `teams_login` or delete these files.

## Development

```bash
npm run dev          # Run MCP server in dev mode
npm run build        # Compile TypeScript
npm run research     # Explore Teams APIs (logs network calls)
npm test             # Run unit tests
npm run typecheck    # TypeScript type checking
```

See [AGENTS.md](AGENTS.md) for detailed architecture and contribution guidelines.

---

## Teams Chat Export Bookmarklet

This repo also includes a standalone bookmarklet for exporting Teams chat messages to Markdown. See [teams-bookmarklet/README.md](teams-bookmarklet/README.md).
