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
├── utils/
│   ├── parsers.ts        # Pure parsing functions (testable)
│   └── parsers.test.ts   # Unit tests for parsers
├── __fixtures__/
│   └── api-responses.ts  # Mock API responses for testing
├── types/
│   └── teams.ts          # TypeScript interfaces
└── test/                 # Integration test tools (CLI, MCP harness)
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

### Authentication Patterns
Different Teams APIs use different authentication mechanisms:

| API | Auth Method | Helper Function |
|-----|-------------|-----------------|
| **Search** (Substrate v2/query) | JWT Bearer token from MSAL | `getValidToken()` |
| **People/Suggestions** (Substrate v1/suggestions) | Same JWT + `cvid`/`logicalId` fields | `getValidToken()` |
| **Messaging** (chatsvc) | `skypetoken_asm` cookie | `extractMessageAuth()` |
| **Favorites** (csa/conversationFolders) | CSA token from MSAL + `skypetoken_asm` | `extractCsaToken()` + `extractMessageAuth()` |
| **Threads** (chatsvc) | `skypetoken_asm` cookie | `extractMessageAuth()` |

**Important**: The CSA API (for favorites) requires a GET request to retrieve data, POST only for modifications. The Substrate suggestions API requires `cvid` and `logicalId` correlation IDs in the request body.

### Session Persistence
Playwright's `storageState()` is used to save and restore browser sessions. This means:
- Session cookies help with faster re-authentication
- MSAL tokens refresh automatically when you perform actions in the browser
- After a browser-based search, tokens are captured and cached for direct API use

## MCP Tools

| Tool | Purpose |
|------|---------|
| `teams_search` | Search messages with query operators, supports pagination |
| `teams_send_message` | Send a message to a Teams conversation |
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_get_frequent_contacts` | Get frequently contacted people (for name resolution) |
| `teams_search_people` | Search for people by name or email |
| `teams_login` | Trigger manual login (visible browser) |
| `teams_status` | Check auth status (search, messaging, favorites tokens) |
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation to favourites |
| `teams_remove_favorite` | Unpin a conversation from favourites |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark from a message |
| `teams_get_thread` | Get messages from a conversation/thread |

### Design Philosophy

The toolset follows a **minimal tool philosophy**: fewer, more powerful tools that AI can compose together. Rather than convenience wrappers for common patterns, the AI builds queries using search operators.

### teams_search Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search query with optional operators |
| from | number | 0 | Starting offset for pagination |
| size | number | 25 | Page size |
| maxResults | number | 25 | Maximum results to return |

**Search Operators:**

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:sarah@company.com` | Messages from a person (use actual email) |
| `sent:` | `sent:today`, `sent:lastweek` | Messages by date |
| `in:` | `in:project-alpha` | Messages in a channel |
| `"Name"` | `"Rob Smith"` | Find @mentions (display name in quotes) |
| `NOT` | `NOT from:user@email.com` | Exclude results |
| `hasattachment:` | `hasattachment:true` | Messages with files |

**⚠️ Common Mistakes - What Does NOT Work:**

| Invalid | Why | Use Instead |
|---------|-----|-------------|
| `@me` | Not a valid Teams operator | Use `teams_get_me` to get email/name, then search with those |
| `from:me` | `me` is not recognised | `from:actual.email@company.com` |
| `to:me` | Not supported | Search for `"Display Name"` to find @mentions |
| `mentions:me` | Not supported | Search for `"Display Name"` to find @mentions |

**Common Patterns:**

1. **Messages FROM me:**
   ```
   # First call teams_get_me to get email, then:
   from:rob.smith@company.com
   ```

2. **Messages mentioning me (@mentions):**
   ```
   # First call teams_get_me to get displayName and email, then:
   "Rob Smith" NOT from:rob.smith@company.com
   ```
   The `NOT from:` excludes your own messages where you might have typed your name.

3. **Messages from a specific person:**
   ```
   # If you know their email:
   from:sarah.jones@company.com
   
   # If you only know their name, first call teams_search_people to find their email
   ```

4. **Unread/unanswered questions to me:**
   ```
   # Search for mentions with question marks:
   "Rob Smith" ? NOT from:rob.smith@company.com
   ```

**Response** includes:
- `results[]` with `id`, `content`, `sender`, `timestamp`, `conversationId`, `messageId`, `messageLink`
- `pagination` object with `from`, `size`, `returned`, `total` (if known), `hasMore`, `nextFrom`

The `conversationId` enables replying to search results via `teams_send_message`.
The `messageLink` is a direct URL to open the message in Teams (format: `https://teams.microsoft.com/l/message/{conversationId}/{timestamp}`).

### teams_send_message Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| content | string | required | Message content (HTML supported) |
| conversationId | string | `48:notes` | Conversation to send to. Default is self-chat (notes). |

**Note:** Messaging uses different authentication than search. It requires session cookies (`skypetoken_asm`, `authtoken`) rather than Bearer tokens. These are automatically extracted from the saved session state.

### teams_get_me Parameters

No parameters. Returns current user's profile including `id`, `mri`, `email`, `displayName`, and `tenantId`.

### teams_get_frequent_contacts Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| limit | number | 50 | Maximum number of contacts to return (1-500) |

**Response** includes:
- `contacts[]` with `id`, `mri`, `displayName`, `email`, `givenName`, `surname`, `jobTitle`, `department`, `companyName`
- `returned` count

**Use case:** When a user refers to someone by first name (e.g., "What's Rob been up to?"), call this tool first to get a ranked list of frequent contacts. Match the name against this list to resolve ambiguity before searching messages.

### teams_search_people Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search term (name, email, or partial match) |
| limit | number | 10 | Maximum number of results (1-50) |

**Response** includes:
- `results[]` with `id`, `mri`, `displayName`, `email`, `jobTitle`, `department`, `companyName`

Use this when searching for a specific person by name or email, rather than getting the user's common contacts.

### teams_get_favorites Parameters

No parameters. Returns list of pinned conversation IDs.

### teams_add_favorite / teams_remove_favorite Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation ID to pin/unpin |

### teams_save_message / teams_unsave_message Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | Conversation containing the message |
| messageId | string | required | The message ID to save/unsave |

**Note:** These tools use the same session cookie authentication as messaging.

### teams_get_thread Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation ID to get messages from |
| limit | number | 50 | Maximum number of messages to return (1-200) |

**Response** includes:
- `conversationId` - The conversation ID
- `messageCount` - Number of messages returned
- `messages[]` with:
  - `id` - Message ID (numeric string)
  - `content` - Message content (HTML stripped)
  - `contentType` - Message type (e.g., "RichText/Html")
  - `sender` - Object with `mri` and `displayName`
  - `timestamp` - ISO timestamp
  - `isFromMe` - Whether message is from the current user
  - `messageLink` - Direct link to open this message in Teams

**Use case:** Check for replies to a specific message, read thread context before replying, or review recent messages in a conversation. Use the `conversationId` from search results.

**Note:** Messages are sorted oldest-first. This uses the same session cookie authentication as messaging.

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

# Search for people
npm run test:mcp -- people "john smith"

# Get favourites
npm run test:mcp -- favorites

# Save/unsave a message
npm run test:mcp -- save --to "conversation-id" --message "message-id"
npm run test:mcp -- unsave --to "conversation-id" --message "message-id"

# Get thread/conversation messages
npm run test:mcp -- thread --to "conversation-id"
npm run test:mcp -- thread --to "conversation-id" --limit 20
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

### Search Doesn't Find All Thread Replies
The Substrate search API is a **full-text search** — it only returns messages matching the search terms. If someone replied to your message but their reply doesn't contain your search keywords, it won't appear in results.

**Example:** Searching for "Easter blockout" won't find a reply that says "Given World of Frozen opens the week before, I'd put a fair amount of money on 'yes'" — even though it's a direct reply.

**Workaround:** After finding a message of interest, use `teams_get_thread` with the `conversationId` to retrieve the full thread context including all replies.

### Message Deep Links
For channel threaded messages, deep links use:
- The thread ID (`ClientThreadId`) — the specific thread within a channel
- The message's own timestamp (`DateTimeReceived`) — the exact message, not the parent

The link format is: `https://teams.microsoft.com/l/message/{threadId}/{messageTimestamp}`

Note: The `conversationId` returned in search results for threaded replies will be the thread ID (e.g., `19:0df465dd...@thread.tacv2`) not the channel ID (e.g., `19:-eGaQP4gB...@thread.tacv2`).

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

## Unit Testing

The project uses Vitest for unit testing pure functions. Tests focus on outcomes, not implementations.

### Running Tests

```bash
npm test              # Run all tests once
npm run test:watch    # Run tests in watch mode
npm run test:coverage # Run tests with coverage report
npm run typecheck     # TypeScript type checking only
```

### Test Structure

- **`src/utils/parsers.ts`**: Pure parsing functions extracted for testability
- **`src/utils/parsers.test.ts`**: Unit tests for all parsing functions
- **`src/__fixtures__/api-responses.ts`**: Mock API response data based on real API structures

### What's Tested

The unit tests cover:
- HTML stripping and entity decoding (`stripHtml`)
- Teams deep link generation (`buildMessageLink`)
- Message timestamp extraction (`extractMessageTimestamp`)
- Person suggestion parsing (`parsePersonSuggestion`)
- Search result parsing (`parseV2Result`, `parseSearchResults`)
- JWT profile extraction (`parseJwtProfile`)
- Token expiry calculations (`calculateTokenStatus`)
- People results parsing (`parsePeopleResults`)

### Adding New Tests

When adding new parsing logic:
1. Add the pure function to `src/utils/parsers.ts`
2. Add fixture data to `src/__fixtures__/api-responses.ts` based on real API responses
3. Write tests in `src/utils/parsers.test.ts` that verify expected outputs

### CI/CD

GitHub Actions runs on every push and PR:
- Type checking (`npm run typecheck`)
- Unit tests (`npm test`)
- Build (`npm run build`)

See `.github/workflows/ci.yml` for the workflow configuration.

## Integration Testing

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
| `teams_search_people` | Substrate suggestions | Easy | ✅ Implemented |
| `teams_get_frequent_contacts` | Substrate peoplecache | Easy | ✅ Implemented |
| `teams_get_favorites` | conversationFolders API | Easy | ✅ Implemented |
| `teams_add_favorite` | conversationFolders API | Easy | ✅ Implemented |
| `teams_remove_favorite` | conversationFolders API | Easy | ✅ Implemented |
| `teams_save_message` | rcmetadata API | Easy | ✅ Implemented |
| `teams_unsave_message` | rcmetadata API | Easy | ✅ Implemented |
| `teams_get_thread` | chatsvc messages API | Easy | ✅ Implemented |
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
- `vitest`: Unit testing framework (dev)
