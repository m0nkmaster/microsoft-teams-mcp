# Teams MCP Server

An MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams via direct API calls.

## How It Works

This server calls Microsoft's internal Teams APIs directly (Substrate, chatsvc, CSA) - the same APIs the Teams web app uses. Authentication is handled by opening a browser for you to log in, then extracting and caching the OAuth tokens.

1. **First use**: Opens browser → you log in → tokens extracted and cached
2. **Normal operation**: Direct API calls using cached tokens (no browser needed)
3. **Token refresh**: When tokens expire (~1 hour), browser opens briefly to refresh them automatically using saved session cookies

## Why This Approach?

Microsoft Graph API requires Azure AD app registration, admin consent, and specific permissions. This project uses your existing Teams credentials instead - no Azure setup required.

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

# Send a message to yourself (notes)
npm run cli -- send "Hello from Teams MCP!"

# Send to specific conversation
npm run cli -- send "Message" --to "conversation-id"
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

### Core Tools

| Tool | Description |
|------|-------------|
| `teams_search` | Search messages with operators (`from:`, `sent:`, `in:`, etc.) |
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_send_message` | Send a message to a Teams conversation (default: self-chat) |
| `teams_login` | Trigger manual login (visible browser) |
| `teams_status` | Check authentication and session state |

### People & Contacts

| Tool | Description |
|------|-------------|
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (for name resolution) |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |

### Conversations & Channels

| Tool | Description |
|------|-------------|
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_find_channel` | Find channels by name (your teams + org-wide) |
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |

### Message Actions

| Tool | Description |
|------|-------------|
| `teams_reply_to_thread` | Reply to a channel message as a threaded reply |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark from a message |

### Search Operators

```
from:sarah@company.com     # Messages from person
sent:today                 # Messages from today
sent:lastweek              # Messages from last week
in:project-alpha           # Messages in channel
"Rob Smith"                # Find @mentions (name in quotes)
hasattachment:true         # Messages with files
NOT from:email@co.com      # Exclude results
```

**Note:** `@me`, `from:me`, `to:me` do NOT work. Use `teams_get_me` first to get your email/displayName, then use those values.

Combine: `from:sarah@co.com sent:lastweek hasattachment:true`

## MCP Resources

The server also exposes passive resources for context discovery:

| Resource URI | Description |
|--------------|-------------|
| `teams://me/profile` | Current user's profile (email, displayName, ID) |
| `teams://me/favorites` | Pinned/favourite conversations |
| `teams://status` | Authentication status for all APIs |

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

- Requires initial manual login via browser
- Uses undocumented Microsoft APIs (may change without notice)
- Token refresh opens browser briefly (~1 hour intervals); manual re-login only needed if session cookies expire

---

## Teams Chat Export Bookmarklet

This repo also includes a standalone bookmarklet for exporting Teams chat messages to Markdown.

See [teams-bookmarklet/README.md](teams-bookmarklet/README.md) for details.
