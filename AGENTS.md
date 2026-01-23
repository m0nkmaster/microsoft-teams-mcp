# Agent Guidelines for Teams MCP

This document captures project knowledge to help AI agents work effectively with this codebase.

## Project Overview

This is an MCP (Model Context Protocol) server that enables AI assistants to search Microsoft Teams messages. Rather than using the complex Microsoft Graph API, it uses Playwright browser automation to interact with the Teams web app directly.

## Architecture

```
src/
├── index.ts              # Entry point, runs the MCP server
├── server.ts             # MCP server with tool definitions
├── browser/
│   ├── context.ts        # Playwright browser/context management
│   ├── session.ts        # Session persistence (cookies, storage, token expiry)
│   └── auth.ts           # Authentication detection and handling
├── teams/
│   ├── direct-api.ts     # Direct HTTP calls to Substrate API (preferred)
│   ├── search.ts         # Browser-based search (fallback)
│   ├── messages.ts       # Message extraction from DOM
│   └── api-interceptor.ts # Network request interception
└── types/
    └── teams.ts          # TypeScript interfaces
```

## Key Design Decisions

### Direct API over Browser Automation
The search implementation uses a hybrid approach:

1. **Direct API (preferred)**: Makes HTTP requests directly to the Substrate v2 search API using extracted authentication tokens. No browser needed after initial login.

2. **Browser Fallback**: If no valid token is available (first run or token expired), opens a visible browser for login, then extracts tokens for future use.

### Authentication Flow
1. **First search**: Opens browser → user logs in → search performed → tokens extracted → browser closed
2. **Subsequent searches**: Uses cached tokens for direct API calls (no browser)
3. **Token expiry**: When tokens expire (~1 hour), falls back to browser to refresh them

### Direct API Details
The Substrate v2 query API (`substrate.office.com/searchservice/api/v2/query`) provides:
- Structured JSON responses with message content, sender info, timestamps
- Offset-based pagination (`from`/`size` parameters)
- Total result counts for accurate pagination
- Hit-highlighted search snippets

### Token Management
- Tokens are extracted from browser localStorage after a successful search
- The Substrate search token (`SubstrateSearch-Internal.ReadWrite` scope) is required for search
- Tokens typically expire after ~1 hour
- Expired tokens trigger automatic browser fallback

### Messaging Authentication
Messaging uses a different auth mechanism than search:
- **Search**: Uses JWT Bearer tokens from MSAL localStorage entries
- **Messaging**: Uses session cookies (`skypetoken_asm`, `authtoken`) from Playwright's `storageState()`

The `extractMessageAuth()` function in `direct-api.ts` extracts these cookies for sending messages without needing an active browser.

### Session Persistence
Playwright's `storageState()` is used to save and restore browser sessions. This means:
- Session cookies help with faster re-authentication
- MSAL tokens refresh automatically when you perform actions in the browser
- After a browser-based search, tokens are captured and cached for direct API use

## MCP Tools

| Tool | Purpose |
|------|---------|
| `teams_search` | Search messages with query, supports pagination (from, size) |
| `teams_send_message` | Send a message to a Teams conversation |
| `teams_login` | Trigger manual login (visible browser) |
| `teams_status` | Check authentication and session state |

### teams_search Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search query |
| from | number | 0 | Starting offset for pagination |
| size | number | 25 | Page size |
| maxResults | number | 25 | Maximum results to return |

Response includes `pagination` object with `from`, `size`, `returned`, `total` (if known), `hasMore`, and `nextFrom`.

### teams_send_message Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| content | string | required | Message content (HTML supported) |
| conversationId | string | `48:notes` | Conversation to send to. Default is self-chat (notes). |

**Note:** Messaging uses different authentication than search. It requires session cookies (`skypetoken_asm`, `authtoken`) rather than Bearer tokens. These are automatically extracted from the saved session state.

## Development Commands

```bash
npm run research      # Explore Teams APIs (visible browser, logs network calls)
npm run dev           # Run MCP server in development mode
npm run build         # Compile TypeScript
npm start             # Run compiled MCP server
```

## Testing Tools

Several CLI tools are available for testing and debugging:

### MCP Protocol Test Harness

Tests the server through the actual MCP protocol using in-memory transports. This verifies the full MCP layer works correctly, not just the underlying functions.

```bash
# List available MCP tools
npm run test:mcp

# Search via MCP protocol
npm run test:mcp -- search "your query"

# Search with pagination
npm run test:mcp -- search "your query" --from 25 --size 10

# Check status via MCP
npm run test:mcp -- status

# Output raw MCP response as JSON
npm run test:mcp -- search "your query" --json

# Send a message to yourself (notes)
npm run test:mcp -- send "Hello from MCP!"

# Send to specific conversation
npm run test:mcp -- send "Message" --to "conversation-id"
```

### Direct CLI Tools

```bash
# Check session status
npm run cli -- status

# Search Teams (visible browser)
npm run cli -- search "your query"

# Search with debug output
npm run cli -- search "your query" --debug

# Search in headless mode (requires saved session)
npm run cli -- search "your query" --headless

# Output as JSON
npm run cli -- search "your query" --json

# Pagination: get page 2 (results 25-49)
npm run cli -- search "your query" --from 25 --size 25

# Send a message to yourself (notes)
npm run cli -- send "Hello from CLI!"

# Send to specific conversation
npm run cli -- send "Message" --to "conversation-id"

# Login flow
npm run cli -- login
npm run cli -- login --force  # Clear session and re-login

# Full test suite
npm run test:manual
npm run test:manual -- --search "your query"

# Debug search with screenshots
npm run debug:search
npm run debug:search -- "your query"
```

The `debug:search` command saves screenshots to `debug-output/` and is useful when selectors need updating.

## Common Issues and Solutions

### Session Expired
If searches fail with authentication errors:
1. Call `teams_login` with `forceNew: true`
2. Or delete `session-state.json` and run `npm run research`

### Search Returns Empty Results
- Teams UI selectors may have changed; check `src/teams/search.ts` for selector updates
- The API interception patterns may need updating; check `src/teams/api-interceptor.ts`

### Browser Won't Launch
- Ensure Playwright browsers are installed: `npx playwright install chromium`
- Check for existing browser processes that may be blocking

## File Locations

- **Session state**: `./session-state.json` (gitignored)
- **Browser profile**: `./.user-data/` (gitignored)
- **Debug output**: `./debug-output/` (gitignored, screenshots and HTML dumps)
- **API research docs**: `./docs/API-RESEARCH.md`

## Extending the MCP

### Adding New Tools
1. Add tool definition to `TOOLS` array in `src/server.ts`
2. Add input schema with Zod in `src/server.ts`
3. Handle the tool in the switch statement in the request handler

### Updating Selectors
Teams may update their UI. Key selector files:
- `src/teams/search.ts`: Search box and result selectors
- `src/browser/auth.ts`: Authentication detection selectors

Reference: `teams-export/teams-export.js` contains a working bookmarklet with proven DOM selectors for Teams message extraction. Key selectors include:
- `[data-tid="chat-pane-item"]` - Message container
- `[data-tid="chat-pane-message"]` - Message body
- `[data-tid="message-author-name"]` - Sender name
- `[id^="content-"]:not([id^="content-control"])` - Message content

### Capturing New API Endpoints
Run `npm run research`, perform actions in Teams, and check the terminal output for captured requests.

## Testing Approach

Due to the nature of browser automation against a live service:
- Use `npm run test:mcp -- search "query"` to test via the full MCP protocol layer
- Use `npm run cli -- search "query" --debug` for quick testing of underlying functions
- Use `npm run debug:search` when selectors need investigation (saves screenshots)
- Use `npm run research` to explore new API patterns (logs all network traffic)
- Check `debug-output/` for screenshots and HTML dumps when debugging

The MCP test harness (`test:mcp`) uses the SDK's `InMemoryTransport` to connect a test client to the server in-process, verifying that tool definitions, input validation, and response formatting all work correctly through the protocol layer.

## Key API Endpoints Discovered

From research, Teams uses these primary APIs:

### Search & Query
| Endpoint | Purpose |
|----------|---------|
| `substrate.office.com/searchservice/api/v2/query` | Full message search with pagination |
| `substrate.office.com/search/api/v1/suggestions` | People/message typeahead |
| `substrate.office.com/search/api/v1/suggestions?scenario=peoplecache` | Frequent contacts list |

### Channels & Messages
| Endpoint | Purpose |
|----------|---------|
| `teams.microsoft.com/api/csa/{region}/api/v1/containers/{id}/posts` | Channel messages |
| `teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{id}/messages` | Send/receive messages |
| `teams.microsoft.com/api/chatsvc/{region}/v1/threads/{id}/annotations` | Reactions, read status |
| `teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders` | Favorites/pinned chats |
| `teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{id}/rcmetadata/{mid}` | Save/unsave messages |

### People & Profile
| Endpoint | Purpose |
|----------|---------|
| `nam.loki.delve.office.com/api/v2/person` | Detailed person profile |
| `nam.loki.delve.office.com/api/v1/schedule` | Working hours, availability |
| `nam.loki.delve.office.com/api/v1/oofstatus` | Out of office status |
| `teams.microsoft.com/api/mt/part/{region}/beta/users/fetch` | Batch user lookup |

### Files & Attachments
| Endpoint | Purpose |
|----------|---------|
| `substrate.office.com/AllFiles/api/users(...)/AllShared` | Files shared in conversation |

Regional identifiers: `amer`, `emea`, `apac`

See `docs/API-RESEARCH.md` for full endpoint documentation with request/response examples.

## Potential Future Tools

Based on API research, these tools could be implemented:

| Tool | API | Difficulty | Status |
|------|-----|------------|--------|
| `teams_get_me` | JWT token extraction | Easy | ✅ Implemented |
| `teams_send_message` | chatsvc messages API | Medium | ✅ Implemented |
| `teams_get_favorites` | conversationFolders API | Easy | Ready |
| `teams_add_favorite` | conversationFolders API | Easy | Ready |
| `teams_save_message` | rcmetadata API | Easy | Ready |
| `teams_search_people` | Substrate suggestions | Easy | Pending |
| `teams_get_person` | Delve person API | Easy | Pending |
| `teams_get_channel_posts` | CSA containers API | Medium | Pending |
| `teams_get_files` | AllFiles API | Medium | Pending |

### Not Yet Feasible
- **Get all saved messages** - No single endpoint; saved flag is per-message in rcMetadata
- **Chat list** - Data loaded at startup, not in separate API
- **Activity feed** - Exists at `48:notifications` but format unclear
- **Presence/Status** - Real-time via WebSocket, not HTTP
- **Calendar** - Outlook APIs exist but need separate research

## Dependencies

- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `playwright`: Browser automation
- `zod`: Runtime input validation
