# Teams MCP Server

An MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams-search messages, send replies, manage favourites, and more.

## How It Works

This server calls Microsoft's internal Teams APIs directly (Substrate, chatsvc, CSA)-the same APIs the Teams web app uses. No Azure AD app registration or admin consent required.

**Authentication flow:**
1. Run `teams_login` to open a browser and log in
2. OAuth tokens are extracted and cached
3. All operations use cached tokens directly (no browser needed)
4. When tokens expire (~1 hour), run `teams_login` again

## Installation

### Prerequisites

- Node.js 18+
- A Microsoft account with Teams access
- Google Chrome, Microsoft Edge, or Chromium browser installed

### Configure Your MCP Client

Add to your MCP client configuration (e.g., Claude Desktop, Cursor):

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "msteams-mcp"]
    }
  }
}
```

That's it. The server uses your system's Chrome (macOS/Linux) or Edge (Windows) for authentication.

### Manual Installation (optional)

If you prefer to install globally:

```bash
npm install -g msteams-mcp
```

Then configure:

```json
{
  "mcpServers": {
    "teams": {
      "command": "msteams-mcp"
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
| `teams_get_activity` | Get activity feed (mentions, reactions, replies, notifications) |

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
| `teams_get_unread` | Get unread counts (aggregate or per-conversation) |
| `teams_mark_read` | Mark a conversation as read up to a message |

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

## CLI Tools (Development)

For local development, CLI tools are available for testing and debugging:

```bash
# Check authentication status
npm run cli -- status

# Search messages
npm run cli -- search "meeting notes"
npm run cli -- search "project" --from 0 --size 50

# Send messages (default: your own notes/self-chat)
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
npm run test:mcp -- activity              # Get activity feed
npm run test:mcp -- unread                # Check unread counts
```

## Limitations

- **Login required** - Run `teams_login` to authenticate (opens browser)
- **Token expiry** - Tokens expire after ~1 hour; run `teams_login` again when needed
- **Undocumented APIs** - Uses Microsoft's internal APIs which may change without notice
- **Search limitations** - Full-text search only; thread replies not matching search terms won't appear (use `teams_get_thread` for full context)
- **Own messages only** - Edit/delete only works on your own messages

## Session Files

These files are created locally and gitignored:

- `session-state.json` - Encrypted browser session
- `token-cache.json` - Encrypted OAuth tokens
- `.user-data/` - Browser profile

If your session expires, call `teams_login` or delete these files.

## Development

For local development:

```bash
git clone https://github.com/m0nkmaster/microsoft-teams-mcp.git
cd microsoft-teams-mcp
npm install
npm run build
```

Development commands:

```bash
npm run dev          # Run MCP server in dev mode
npm run build        # Compile TypeScript
npm run research     # Explore Teams APIs (logs network calls)
npm test             # Run unit tests
npm run typecheck    # TypeScript type checking
```

For development with hot reload, configure your MCP client:

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["tsx", "/path/to/microsoft-teams-mcp/src/index.ts"]
    }
  }
}
```

See [AGENTS.md](AGENTS.md) for detailed architecture and contribution guidelines.

---

## Teams Chat Export Bookmarklet

This repo also includes a standalone bookmarklet for exporting Teams chat messages to Markdown. See [teams-bookmarklet/README.md](teams-bookmarklet/README.md).
