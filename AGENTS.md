# Agent Guidelines for Teams MCP

This document captures project knowledge to help AI agents work effectively with this codebase.

## Package Information

- **npm package**: `msteams-mcp`
- **Repository**: https://github.com/m0nkmaster/microsoft-teams-mcp
- **Install**: `npx -y msteams-mcp@latest` or `npm install -g msteams-mcp`

### Publishing Updates

To publish a new version:

```bash
npm run build                    # Compile TypeScript
npm version patch|minor|major    # Bump version
npm publish --otp=XXXXXX         # Publish (requires 2FA OTP)
git push && git push --tags      # Push version commit and tag
```

The `files` field in `package.json` limits the published package to `dist/` and `README.md` only.

## Project Overview

This is an MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams. Rather than using the complex Microsoft Graph API, it uses Teams APIs (Substrate, chatsvc, CSA) with authentication tokens extracted from a browser session. The browser is only used for initial login - all operations use direct API calls.

## Architecture

### Directory Structure

```
src/
├── index.ts              # Entry point, runs the MCP server
├── server.ts             # MCP server (TeamsServer class) - delegates to tool registry
├── constants.ts          # Shared constants (page sizes, timeouts, thresholds)
├── tools/                # Tool handlers (modular design)
│   ├── index.ts          # Tool context and type definitions
│   ├── registry.ts       # Tool registry - maps names to handlers
│   ├── search-tools.ts   # Search and channel tools
│   ├── message-tools.ts  # Messaging, favourites, save/unsave tools
│   ├── people-tools.ts   # People search and profile tools
│   └── auth-tools.ts     # Login and status tools
├── auth/                 # Authentication and credential management
│   ├── index.ts          # Module exports
│   ├── crypto.ts         # AES-256-GCM encryption for credentials at rest
│   ├── session-store.ts  # Secure session state storage with encryption
│   ├── token-extractor.ts # Extract tokens from Playwright session state
│   └── token-refresh.ts  # Proactive token refresh via OAuth2 endpoint
├── api/                  # API client modules (one per API surface)
│   ├── index.ts          # Module exports
│   ├── substrate-api.ts  # Search and people APIs (Substrate v2)
│   ├── chatsvc-api.ts    # Messaging, threads, save/unsave (chatsvc)
│   └── csa-api.ts        # Favorites API (CSA)
├── browser/              # Playwright browser automation (login only)
│   ├── context.ts        # Browser/context management with encrypted session
│   └── auth.ts           # Authentication detection and manual login handling
├── utils/
│   ├── parsers.ts        # Pure parsing functions (testable)
│   ├── parsers.test.ts   # Unit tests for parsers
│   ├── http.ts           # HTTP client with retry, timeout, error handling
│   ├── api-config.ts     # API endpoints and header configuration
│   └── auth-guards.ts    # Reusable auth check utilities (Result types)
├── types/
│   ├── teams.ts          # Teams data interfaces
│   ├── errors.ts         # Error taxonomy with machine-readable codes
│   └── result.ts         # Result<T, E> type for explicit error handling
├── __fixtures__/
│   └── api-responses.ts  # Mock API responses for testing
└── test/                 # Integration test tools (CLI, MCP harness)
```

### Implementation Patterns

1. **Credential Encryption**: Session state and token cache are encrypted at rest using AES-256-GCM with a machine-specific key derived from hostname and username. Files have restrictive permissions (0o600).

2. **Server Class Pattern**: `TeamsServer` class encapsulates all state (browser manager, initialisation flag), allowing multiple server instances and simpler testing.

3. **Error Taxonomy**: Errors use machine-readable codes (`ErrorCode` enum), `retryable` flags, and `suggestions` arrays to help LLMs understand failures and recover appropriately.

4. **Result Types**: API functions return `Result<T, McpError>` for type-safe error handling with explicit success/failure discrimination.

5. **HTTP Utilities**: Centralised HTTP client (`utils/http.ts`) provides automatic retry with exponential backoff, request timeouts, and rate limit tracking.

6. **MCP Resources**: Passive resources (`teams://me/profile`, `teams://me/favorites`, `teams://status`) provide context discovery without tool calls.

7. **Tool Registry Pattern**: Tools are organised into logical groups (`search-tools.ts`, `message-tools.ts`, etc.) with a central registry (`tools/registry.ts`). This enables:
   - Better separation of concerns
   - Easier testing of individual tools
   - Simpler addition of new tools

8. **Auth Guards**: Reusable authentication check utilities in `utils/auth-guards.ts` return `Result` types for consistent error handling across API modules.

9. **Shared Constants**: Magic numbers are centralised in `constants.ts` for maintainability (page sizes, timeouts, thresholds).

## How It Works

### Authentication Flow

All operations use direct API calls to Teams APIs. The browser is only used for authentication:

1. **Login**: Opens visible browser → user authenticates → session state saved → browser closed
2. **All subsequent operations**: Use cached tokens for direct API calls (no browser)
3. **Token expiry**: When tokens expire (~1 hour), proactive refresh is attempted; if that fails, user must re-authenticate via `teams_login`

This approach provides faster, more reliable operations compared to DOM scraping, with structured JSON responses and proper pagination support.

The server uses the system's installed browser rather than downloading Playwright's bundled Chromium (~180MB savings):

- **Windows**: Uses Microsoft Edge (always pre-installed on Windows 10+)
- **macOS/Linux**: Uses Google Chrome

This is configured via Playwright's `channel` option in `src/browser/context.ts`. If the system browser isn't available, a helpful error message suggests installing Chrome or running `npx playwright install chromium` as a fallback.

### Token Management

- Tokens are extracted from browser localStorage after login
- The Substrate search token (`SubstrateSearch-Internal.ReadWrite` scope) is required for search
- Tokens typically expire after ~1 hour
- **Proactive token refresh**: When tokens have less than 10 minutes remaining, the server automatically refreshes them using a headless browser. MSAL handles the token refresh when Teams loads, then we save the updated session state.
- This is seamless to the user - the browser is invisible
- If refresh fails, user must re-authenticate via `teams_login`

**How token refresh works:**
1. `requireSubstrateTokenAsync()` checks if tokens are expired or have <10 minutes remaining
2. If so, `refreshTokensViaBrowser()` opens a headless browser with saved session
3. Navigates to Teams and triggers a search (MSAL only refreshes tokens when an API call requires them)
4. The search triggers MSAL's `acquireTokenSilent` which refreshes the Substrate token
5. Saves the updated session state with new tokens
6. Subsequent API calls use the refreshed tokens

**Important:** MSAL doesn't automatically refresh tokens on page load - it only acquires new tokens when an API call actually needs them. Simply loading Teams isn't enough; we must trigger a search to force token acquisition.

**Testing token refresh:**
```bash
npm run cli -- refresh
```

### API Authentication

Different Teams APIs use different authentication mechanisms:

| API | Auth Method | Module | Helper Function |
|-----|-------------|--------|-----------------|
| **Search** (Substrate v2/query) | JWT Bearer token from MSAL | `auth/token-extractor` | `getValidSubstrateToken()` |
| **People/Suggestions** (Substrate v1/suggestions) | Same JWT + `cvid`/`logicalId` fields | `auth/token-extractor` | `getValidSubstrateToken()` |
| **Messaging** (chatsvc) | `skypetoken_asm` cookie | `auth/token-extractor` | `extractMessageAuth()` |
| **Favorites** (csa/conversationFolders) | CSA token from MSAL + `skypetoken_asm` | `auth/token-extractor` | `extractCsaToken()` + `extractMessageAuth()` |
| **Threads** (chatsvc) | `skypetoken_asm` cookie | `auth/token-extractor` | `extractMessageAuth()` |

**Important**: The CSA API (for favorites) requires a GET request to retrieve data, POST only for modifications. The Substrate suggestions API requires `cvid` and `logicalId` correlation IDs in the request body.

### Session Persistence

Playwright's `storageState()` is used to save browser session state after login. This includes:
- Session cookies (for messaging APIs)
- MSAL tokens in localStorage (for search and people APIs)
- Tokens are extracted and cached for direct API use

Session state and token cache files are protected by:
1. **Encryption at rest**: AES-256-GCM encryption using a key derived from machine-specific values (hostname + username)
2. **File permissions**: Restrictive 0o600 permissions (owner read/write only)
3. **Automatic migration**: Existing plaintext files are automatically encrypted on first read

### Conversation Types

The chatsvc conversation API returns `threadProperties` with type information:

| Type | `threadType` | `productThreadType` | Notes |
|------|--------------|---------------------|-------|
| Standard Channel | `topic` | `TeamsStandardChannel` | Has `groupId`, name in `topicThreadTopic` |
| Team (General) | `space` | `TeamsTeam` | Team root, name in `spaceThreadTopic` |
| Private Channel | `space` | `TeamsPrivateChannel` | Has `groupId`, name in `topicThreadTopic` |
| Meeting Chat | `meeting` | `Meeting` | Name in `topic` |
| Group Chat | `chat` | `Chat` | Name in `topic` or from members |
| 1:1 Chat | `chat` | `OneOnOne` | Name from other participant |

**Name sources:**
- `topicThreadTopic`: Channel name (for channels within a team)
- `spaceThreadTopic`: Team name (for team root conversations)
- `topic`: Meeting title or user-set chat topic
- For chats without topics: extract from `members` array or recent messages

### Virtual Conversations

Teams uses special "virtual conversation" IDs that aggregate data across all conversations. These use the standard messages endpoint but return consolidated views:

| Virtual ID | Purpose | Constant |
|------------|---------|----------|
| `48:saved` | Saved/bookmarked messages | - |
| `48:threads` | Followed threads | - |
| `48:mentions` | @mentions | - |
| `48:notifications` | Activity feed | `NOTIFICATIONS_ID` |
| `48:notes` | Personal notes/self-chat | `SELF_CHAT_ID` |

**Endpoint pattern:** `GET /api/chatsvc/{region}/v1/users/ME/conversations/{virtualId}/messages`

Each message in the response includes a `clumpId` field containing the original conversation ID where the message lives, enabling navigation back to the source.

See `docs/API-REFERENCE.md` for full response structure.

### User ID Formats

Teams APIs return user IDs in multiple formats. The `extractObjectId()` function in `parsers.ts` handles all of these:

| Format | Example | Notes |
|--------|---------|-------|
| Raw GUID | `ab76f827-27e2-4c67-a765-f1a53145fa24` | Standard format |
| MRI | `8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24` | Teams internal identifier |
| ID with tenant | `ab76f827-...@56b731a8-...` | GUID followed by tenant ID |
| Base64-encoded GUID | `93qkaTtFGWpUHjyRafgdhg==` | 16 bytes, little-endian |
| Skype ID | `orgid:ab76f827-...` | Used in skypetoken claims |

**Base64 GUID decoding:** The Substrate people search API returns user IDs as base64-encoded GUIDs. These are 16 bytes encoded in base64 (24 chars with padding). Microsoft uses little-endian byte ordering for the first three GUID groups (Data1, Data2, Data3).

**1:1 Chat ID format:** Conversation IDs for 1:1 chats follow this predictable format:
```
19:{userId1}_{userId2}@unq.gbl.spaces
```
The two user object IDs (GUIDs) are sorted lexicographically. This format works for internal users. External/guest users may require a different format (not researched).

## MCP Tools

### Overview

| Tool | Purpose |
|------|---------|
| `teams_search` | Search messages with query operators, supports pagination |
| `teams_send_message` | Send a message to a Teams conversation |
| `teams_reply_to_thread` | Reply to a channel message as a threaded reply |
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
| `teams_get_saved_messages` | Get list of saved/bookmarked messages with source references |
| `teams_get_followed_threads` | Get list of followed threads with source references |
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_find_channel` | Find channels by name (your teams + org-wide), shows membership |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |
| `teams_edit_message` | Edit one of your own messages |
| `teams_delete_message` | Delete one of your own messages (soft delete) |
| `teams_get_unread` | Get unread status for favourites (aggregate) or specific conversation |
| `teams_mark_read` | Mark a conversation as read up to a specific message |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies, notifications) |
| `teams_search_emoji` | Search for emojis by name (standard + custom org emojis) |
| `teams_add_reaction` | Add an emoji reaction to a message |
| `teams_remove_reaction` | Remove an emoji reaction from a message |

### Design Philosophy

The toolset follows a **minimal tool philosophy**: fewer, more powerful tools that AI can compose together. Rather than convenience wrappers for common patterns, the AI builds queries using search operators.

### Tool Reference

#### teams_search

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search query with optional operators |
| from | number | 0 | Starting offset for pagination |
| size | number | 25 | Page size |
| maxResults | number | 25 | Maximum results to return |

**Search Operators:**

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:sarah@company.com` | Messages from a person (email or name) |
| `to:` | `to:rob macdonald` | Messages to a person (use spaces, not dots/email) |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | Messages by date (use explicit dates) |
| `in:` | `budget in:EEC Leads` | Channel filter - only works reliably WITH content terms (no quotes!) |
| `sent:today` | `sent:today` | Messages from today |
| `"Name"` | `"Rob Smith"` | Find @mentions (display name in quotes) |
| `NOT` | `NOT from:user@email.com` | Exclude results |
| `hasattachment:` | `hasattachment:true` | Messages with files |
| `is:` | `is:Messages`, `is:Meetings`, `is:Channels`, `is:Chats` | Filter by type (case-sensitive, plural) |

**Note:** Results are sorted by recency, so date filters are often unnecessary.

**⚠️ Common Mistakes - What Does NOT Work:**

| Invalid | Why | Use Instead |
|---------|-----|-------------|
| `@me` | Not a valid Teams operator | Use `teams_get_me` to get email/name, then search with those |
| `from:me` | `me` is not recognised | `from:actual.email@company.com` |
| `to:rob.macdonald` | Email format falls back to text search | Use `to:rob macdonald` (spaces, not dots) |
| `mentions:` | Not a valid operator | Search for `"Display Name"` to find @mentions |
| `is:meeting` | Must be plural with capital | Use `is:Meetings` (case-sensitive) |
| `is:Group Chats` | Spaces break it | Use `is:Chats` (no "Group" variant exists) |
| `sent:lastweek` | Not supported by Teams API | Use `sent:>=2026-01-18` or omit (results sorted by recency) |
| `in:EEC Leads` alone | Unreliable without content | Use `content in:EEC Leads` or `teams_get_thread` |
| `sent:thisweek` | Not supported | Use date range like `sent:>=2026-01-20` |
| `in:"EEC Leads"` | Quotes break the operator | `in:EEC Leads` (no quotes, full channel name) |

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

#### teams_send_message

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| content | string | required | Message content. Supports inline @mentions using `@[Name](mri)` syntax. |
| conversationId | string | `48:notes` | Conversation to send to. Default is self-chat (notes). |
| replyToMessageId | string | - | For channel thread replies: the message ID of the thread root. |

**@Mentions:**

Use inline `@[DisplayName](mri)` syntax in the content. Get MRI from `teams_search_people` or `teams_get_frequent_contacts`.

```
teams_send_message content="Hey @[John Smith](8:orgid:abc...), can you review this?"
```

The display name can be any text (e.g., first name only). The MRI determines who gets notified.

**Thread Reply Semantics:**

Teams has different messaging models for channels vs chats:

- **Channels** have threaded conversations. Each top-level post creates a thread, and replies go to that specific thread.
- **Chats** (1:1, group, meeting) are flat - all messages go to the same conversation without threading.

To reply to a channel thread:
1. Get the `conversationId` (channel ID) and `messageId` (thread root ID) from search results or `teams_get_thread`
2. Call `teams_send_message` with both `conversationId` AND `replyToMessageId`

To post a new top-level message in a channel:
- Only provide `conversationId` (the channel ID), omit `replyToMessageId`

**Examples:**

```
# Reply to an existing thread in a channel
teams_send_message content="Thanks!" conversationId="19:channel@thread.tacv2" replyToMessageId="1769274474340"

# Post a new top-level message in a channel
teams_send_message content="Hello team!" conversationId="19:channel@thread.tacv2"

# Send a message in a chat (no threading)
teams_send_message content="Hey!" conversationId="19:abc_def@unq.gbl.spaces"
```

**Response fields:**

| Field | Description |
|-------|-------------|
| `messageId` | Client-generated ID (not useful for subsequent operations) |
| `timestamp` | Server timestamp in milliseconds |
| `serverMessageId` | **Use this for reactions, edits, threading, etc.** (timestamp as string) |
| `conversationId` | The conversation the message was sent to |

**Important:** When performing operations on a newly-sent message (reactions, edits, threading), use `serverMessageId` - not `messageId`. The `messageId` is client-generated and won't work with Teams APIs. The `serverMessageId` is the server-assigned timestamp-based ID.

**Note:** Messaging uses different authentication than search. It requires session cookies (`skypetoken_asm`, `authtoken`) rather than Bearer tokens. These are automatically extracted from the saved session state.

#### teams_reply_to_thread

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| content | string | required | The reply content. Supports inline @mentions using `@[Name](mri)` syntax. |
| conversationId | string | required | The channel conversation ID (from search results). |
| messageId | string | required | The message ID to reply to (from search results). |

**How it works:**

The tool uses the provided `messageId` directly as the thread root. In Teams channels:
- If the message is a top-level post, the reply appears as a threaded reply under that post
- If the message is already a reply within a thread, the reply goes to the same thread

**Important:** The `messageId` from search results is the timestamp-based ID (e.g., `1737445069907`) that Teams uses for threading. This is extracted automatically from search results.

**Example workflow:**

```
1. teams_search "budget report" → returns { conversationId: "19:abc@thread.tacv2", messageId: "1737445069907" }
2. teams_reply_to_thread content="Thanks!" conversationId="19:abc@thread.tacv2" messageId="1737445069907"
```

**Response** includes:
- `messageId` - Your new reply's message ID
- `threadRootMessageId` - The message ID used for the reply
- `conversationId` - The channel ID

#### teams_get_me

No parameters. Returns current user's profile including `id`, `mri`, `email`, `displayName`, and `tenantId`.

#### teams_get_frequent_contacts

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| limit | number | 50 | Maximum number of contacts to return (1-500) |

**Response** includes:
- `contacts[]` with `id`, `mri`, `displayName`, `email`, `givenName`, `surname`, `jobTitle`, `department`, `companyName`
- `returned` count

**Use case:** When a user refers to someone by first name (e.g., "What's Rob been up to?"), call this tool first to get a ranked list of frequent contacts. Match the name against this list to resolve ambiguity before searching messages.

#### teams_search_people

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search term (name, email, or partial match) |
| limit | number | 10 | Maximum number of results (1-50) |

**Response** includes:
- `results[]` with `id`, `mri`, `displayName`, `email`, `jobTitle`, `department`, `companyName`

Use this when searching for a specific person by name or email, rather than getting the user's common contacts.

#### teams_get_favorites

No parameters.

**Response** includes:
- `favorites[]` with `conversationId`, `displayName`, `conversationType`
  - `displayName`: Human-readable name (channel name, chat topic, meeting title, or participant names)
  - `conversationType`: One of `Channel`, `Chat`, or `Meeting`

Name sources by type:
- **Channels**: Channel name from Teams API (e.g., "Support")
- **Meetings**: Meeting title/subject
- **Chats with topic**: The user-set chat topic
- **Chats without topic**: Participant names extracted from recent messages (e.g., "Smith, John, Jones, Sarah + 2 more")

#### teams_add_favorite / teams_remove_favorite

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation ID to pin/unpin |

#### teams_save_message / teams_unsave_message

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | Conversation containing the message |
| messageId | string | required | The message ID to save/unsave |

**Note:** These tools use the same session cookie authentication as messaging.

#### teams_get_saved_messages

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| limit | number | 50 | Maximum number of saved messages to return (1-200) |

**Response** includes:
- `count` - Number of saved messages returned
- `messages[]` with:
  - `content` - Message content (may be empty for some items - use `teams_get_thread` to fetch full content)
  - `sender` - Object with `mri` and `displayName`
  - `timestamp` - ISO timestamp when the message was saved
  - `sourceConversationId` - The original conversation where this message lives
  - `sourceMessageId` - The original message ID in the source conversation
  - `messageLink` - Direct link to open this message in Teams

**Use case:** List bookmarked messages. Use `sourceConversationId` with `teams_get_thread` to retrieve the full message content and context.

**Note:** Returns references to saved messages. Some items may have empty content - these are bookmark pointers. The Teams client fetches actual content separately.

#### teams_get_followed_threads

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| limit | number | 50 | Maximum number of followed threads to return (1-200) |

**Response** includes:
- `count` - Number of followed threads returned
- `threads[]` with:
  - `content` - Thread content preview (may be empty for some items)
  - `sender` - Object with `mri` and `displayName`
  - `timestamp` - ISO timestamp
  - `sourceConversationId` - The original conversation/channel where this thread lives
  - `sourcePostId` - The root post ID of the thread
  - `messageLink` - Direct link to open this thread in Teams

**Use case:** List threads you're following. Use `sourceConversationId` with `teams_get_thread` to retrieve the full thread content.

#### teams_get_thread

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation ID to get messages from |
| limit | number | 50 | Maximum number of messages to return (1-200) |
| markRead | boolean | false | If true, marks the conversation as read up to the latest message after fetching |

**Response** includes:
- `conversationId` - The conversation ID
- `messageCount` - Number of messages returned
- `unreadCount` - Number of unread messages (from others) in the fetched range
- `lastReadMessageId` - The message ID of your last read position
- `markedAsRead` - Whether the conversation was marked as read (only present if `markRead` was true)
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

#### teams_find_channel

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Channel name to search for (partial match) |
| limit | number | 10 | Maximum number of results (1-50) |

**Response** includes:
- `count` - Number of matching channels
- `channels[]` with:
  - `channelId` - Channel conversation ID (use with `teams_get_thread` to read messages)
  - `channelName` - Channel display name
  - `teamName` - Parent team display name
  - `teamId` - Parent team group ID (may be empty for some channels)
  - `channelType` - "Standard", "Private", or "Shared"
  - `description` - Channel description (if set)
  - `isMember` - Whether you're a member of this channel's team

**How it works:** This tool combines two searches:
1. **User's teams/channels** (Teams List API) - Reliable, complete list of channels you're a member of
2. **Organisation-wide discovery** (Substrate suggestions API) - Broader but less reliable typeahead search

Results are merged and deduplicated. Channels from your teams appear first with `isMember: true`.

**Use cases:**
1. Finding channels you're a member of (reliable, includes private channels)
2. Discovering channels to join or follow (org-wide search)
3. Getting channel IDs to read messages with `teams_get_thread`

**Note:** The org-wide search uses a typeahead/autocomplete API which may not find all channels, especially with multi-word queries. Your own team channels are searched reliably via client-side filtering.

#### teams_get_chat

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| userId | string | required | The user's identifier (MRI, object ID with tenant, or raw GUID) |

**Response** includes:
- `conversationId` - The 1:1 conversation ID (use with `teams_send_message`)
- `otherUserId` - The other user's object ID
- `currentUserId` - Your object ID

**Use case:** Get the conversation ID for a 1:1 chat with someone, then use it to send a message. The conversation is automatically created when the first message is sent.

**Example flow:**
```
1. teams_search_people "John Smith" → returns { id: "abc123-..." }
2. teams_get_chat "abc123-..." → returns { conversationId: "19:abc123..._def456...@unq.gbl.spaces" }
3. teams_send_message content="Hello!" conversationId="19:abc123..._def456...@unq.gbl.spaces"
```

**Technical note:** The conversation ID format for 1:1 chats is `19:{id1}_{id2}@unq.gbl.spaces` where the two user object IDs are sorted lexicographically. This is a predictable format - Teams creates the conversation implicitly when the first message is sent.

#### teams_edit_message

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation containing the message |
| messageId | string | required | The message ID to edit (numeric string) |
| content | string | required | The new content for the message |

**Response** includes:
- `message` - Success confirmation
- `conversationId` - The conversation ID
- `messageId` - The edited message ID

**Constraints:** You can only edit your own messages. The API returns 403 Forbidden if you try to edit someone else's message.

**Example:**
```
1. teams_get_thread --to "19:abc@thread.tacv2" → find your message with id "1769276832046"
2. teams_edit_message conversationId="19:abc@thread.tacv2" messageId="1769276832046" content="Updated text"
```

#### teams_delete_message

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation containing the message |
| messageId | string | required | The message ID to delete (numeric string) |

**Response** includes:
- `message` - Success confirmation
- `conversationId` - The conversation ID
- `messageId` - The deleted message ID

**Constraints:**
- You can only delete your own messages
- Channel owners/moderators can delete other users' messages in their channels
- This is a soft delete - the message remains but content becomes empty

**Example:**
```
teams_delete_message conversationId="19:abc@thread.tacv2" messageId="1769276832046"
```

#### teams_get_unread

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | optional | Specific conversation to check. If omitted, checks all favourites. |

**Response (single conversation):**
- `conversationId` - The conversation ID
- `unreadCount` - Number of unread messages
- `lastReadMessageId` - Your last read position
- `latestMessageId` - The most recent message ID

**Response (aggregate mode - no conversationId):**
- `totalUnread` - Total unread messages across all checked favourites
- `conversationsWithUnread` - Number of conversations with unread messages
- `conversations[]` - Array of conversations with unread messages, each with:
  - `conversationId`
  - `displayName`
  - `conversationType`
  - `unreadCount`
- `checked` - Number of favourites checked
- `totalFavorites` - Total number of favourites

**Note:** Aggregate mode checks up to 20 favourites to avoid timeout. Uses the consumption horizon API to determine read position.

**Example:**
```
# Check all favourites
teams_get_unread

# Check specific conversation
teams_get_unread conversationId="19:abc@thread.tacv2"
```

#### teams_mark_read

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation to mark as read |
| messageId | string | required | The message ID to mark as read up to |

**Response** includes:
- `message` - Success confirmation
- `conversationId` - The conversation ID
- `markedUpTo` - The message ID marked as read

**Note:** This marks all messages up to and including the specified message as read. Use with the latest message ID to mark the entire conversation as read.

**Example:**
```
teams_mark_read conversationId="19:abc@thread.tacv2" messageId="1769276832046"
```

#### teams_get_activity

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| limit | number | 50 | Maximum number of activity items to return (1-200) |

**Response** includes:
- `count` - Number of activity items returned
- `activities[]` - Array of activity items, each with:
  - `id` - Activity/message ID
  - `type` - Activity type: `mention`, `reaction`, `reply`, `message`, or `unknown`
  - `content` - Activity content (HTML stripped)
  - `contentType` - Raw message type from API
  - `sender` - Object with `mri` and `displayName`
  - `timestamp` - ISO timestamp
  - `conversationId` - Source conversation where activity occurred
  - `topic` - Conversation/thread topic name (if available)
  - `activityLink` - Direct link to open the activity in Teams
- `syncState` - State token for incremental polling (advanced usage)

**Use case:** Check what's happening - who mentioned you, reacted to your messages, or replied to threads you're in. This is the programmatic equivalent of the Activity tab in Teams.

**Example:**
```
# Get recent activity
teams_get_activity

# Get more activity items
teams_get_activity limit=100
```

#### teams_search_emoji

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| query | string | required | Search term (e.g., "thumbs", "heart", "cat") |

**Response** includes:
- `count` - Number of matching emojis
- `emojis[]` - Array of matches, each with:
  - `key` - The emoji key to use with `teams_add_reaction`
  - `description` - Human-readable description with emoji character
  - `type` - Either `standard` (built-in Teams emoji) or `custom` (org-specific)
  - `category` - For standard emojis: reaction, expression, affection, action, animal, object, other
  - `shortcut` - For custom emojis: the shortcut name

**Quick Reaction Keys (no search needed):**
- `like` (👍), `heart` (❤️), `laugh` (😂), `surprised` (😮), `sad` (😢), `angry` (😠)

#### teams_add_reaction

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation containing the message |
| messageId | string | required | The server-assigned message ID (numeric timestamp string) |
| emoji | string | required | The emoji key (e.g., "like", "heart") |

**Response** includes confirmation of the reaction added.

**Where to get the messageId:**
- From `teams_send_message`: Use the `serverMessageId` field (NOT `messageId`)
- From `teams_get_thread`: Use the `id` field from the message
- From `teams_search`: Use the `messageId` field (already the correct format)

**Example:**
```
teams_add_reaction conversationId="19:abc@thread.tacv2" messageId="1769276832046" emoji="like"
```

**Common mistake:** Using the `messageId` from `teams_send_message` will fail - that's a client-generated ID. Always use `serverMessageId` instead.

#### teams_remove_reaction

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| conversationId | string | required | The conversation containing the message |
| messageId | string | required | The server-assigned message ID (same as `teams_add_reaction`) |
| emoji | string | required | The emoji key to remove |

#### teams_login

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| forceNew | boolean | false | If true, clears existing session and forces fresh login |

Opens a visible browser window for the user to authenticate with Microsoft. After successful login, the session is saved for subsequent API calls.

#### teams_status

No parameters. Returns authentication status for all APIs:
- `directApi` - Substrate search token status (available, expiry time, minutes remaining)
- `messaging` - Whether messaging cookies are available
- `favorites` - Whether favorites API auth is available
- `session` - Whether session file exists and if it's likely expired

## Development

### Commands

```bash
npm run research      # Explore Teams APIs (visible browser, logs network calls)
npm run dev           # Run MCP server in development mode
npm run build         # Compile TypeScript
npm run lint          # Run ESLint (also lint:fix to auto-fix)
npm start             # Run compiled MCP server
```

### Testing

#### MCP Protocol Test Harness

Tests the server through the actual MCP protocol using in-memory transports. This verifies the full MCP layer works correctly, not just the underlying functions.

The harness can call **any tool** generically. Unrecognised commands are treated as tool names (with `teams_` prefix added if missing). Use `--key value` for parameters.

```bash
# List available MCP tools and shortcuts
npm run test:mcp

# Generic tool call (any tool works)
npm run test:mcp -- teams_find_channel --query "support"
npm run test:mcp -- find_channel --query "support"   # auto-prefixes teams_

# Shortcuts for common tools
npm run test:mcp -- search "your query"              # teams_search
npm run test:mcp -- search "your query" --from 25 --size 10
npm run test:mcp -- status                           # teams_status
npm run test:mcp -- me                               # teams_get_me
npm run test:mcp -- login                            # teams_login
npm run test:mcp -- send "Hello from MCP!"           # teams_send_message
npm run test:mcp -- send "Message" --to "conv-id"
npm run test:mcp -- reply "Thanks!" --to "channel-id" --message "msg-id"
npm run test:mcp -- people "john smith"              # teams_search_people
npm run test:mcp -- favorites                        # teams_get_favorites
npm run test:mcp -- contacts                         # teams_get_frequent_contacts
npm run test:mcp -- channel "project-alpha"          # teams_find_channel
npm run test:mcp -- chat "user-guid-or-mri"          # teams_get_chat
npm run test:mcp -- thread --to "conv-id"            # teams_get_thread
npm run test:mcp -- save --to "conv-id" --message "msg-id"
npm run test:mcp -- unsave --to "conv-id" --message "msg-id"
npm run test:mcp -- unread                           # teams_get_unread (aggregate)
npm run test:mcp -- unread --to "conv-id"            # teams_get_unread (specific)
npm run test:mcp -- markread --to "conv-id" --message "msg-id"
npm run test:mcp -- thread --to "conv-id" --markRead
npm run test:mcp -- activity                         # teams_get_activity
npm run test:mcp -- activity --limit 10
npm run test:mcp -- teams_search_emoji --query "heart"    # teams_search_emoji
npm run test:mcp -- teams_add_reaction --conversationId "conv-id" --messageId "msg-id" --emoji "like"
npm run test:mcp -- teams_remove_reaction --conversationId "conv-id" --messageId "msg-id" --emoji "like"

# Output raw MCP response as JSON
npm run test:mcp -- search "your query" --json
```

#### Direct CLI Tools

```bash
# Check session status
npm run cli -- status

# Search Teams (requires valid token - run login first)
npm run cli -- search "your query"

# Output as JSON
npm run cli -- search "your query" --json

# Pagination: get page 2 (results 25-49)
npm run cli -- search "your query" --from 25 --size 25

# Send a message to yourself (notes)
npm run cli -- send "Hello from CLI!"

# Send to specific conversation
npm run cli -- send "Message" --to "conversation-id"

# Login flow (opens browser for authentication)
npm run cli -- login
npm run cli -- login --force  # Clear session and re-login
```

#### Unit Tests

The project uses Vitest for unit testing pure functions. Tests focus on outcomes, not implementations.

```bash
npm test              # Run all tests once
npm run test:watch    # Run tests in watch mode
npm run test:coverage # Run tests with coverage report
npm run typecheck     # TypeScript type checking only
```

**Test Structure:**
- **`src/utils/parsers.ts`**: Pure parsing functions extracted for testability
- **`src/utils/parsers.test.ts`**: Unit tests for all parsing functions
- **`src/__fixtures__/api-responses.ts`**: Mock API response data based on real API structures

**What's Tested:**
- HTML stripping and entity decoding (`stripHtml`)
- Teams deep link generation (`buildMessageLink`)
- Message timestamp extraction (`extractMessageTimestamp`)
- Person suggestion parsing (`parsePersonSuggestion`)
- Search result parsing (`parseV2Result`, `parseSearchResults`)
- JWT profile extraction (`parseJwtProfile`)
- Token expiry calculations (`calculateTokenStatus`)
- People results parsing (`parsePeopleResults`)
- Base64 GUID decoding (`decodeBase64Guid`)
- User ID extraction from various formats (`extractObjectId`)

#### Integration Testing

For testing against the live Teams APIs:
- Use `npm run test:mcp -- search "query"` to test via the full MCP protocol layer
- Use `npm run cli -- search "query"` for quick testing of underlying functions
- Use `npm run research` to explore new API patterns (logs all network traffic)

The MCP test harness (`test:mcp`) uses the SDK's `InMemoryTransport` to connect a test client to the server in-process, verifying that tool definitions, input validation, and response formatting all work correctly through the protocol layer.

#### CI/CD

GitHub Actions runs on every push and PR:
- Linting (`npm run lint`)
- Type checking (`npm run typecheck`)
- Unit tests (`npm test`)
- Build (`npm run build`)
- Documentation review (on main commits, debounced to 2 hours) - checks README/AGENTS.md accuracy against code

See `.github/workflows/ci.yml` and `.github/workflows/doc-reviewer.yml` for workflow configurations.

### Extending the MCP

#### Adding New Tools

1. Choose the appropriate tool file in `src/tools/` (or create a new one for a new category)
2. Define the input schema with Zod
3. Define the tool definition (MCP Tool interface)
4. Implement the handler function returning `ToolResult`
5. Export the registered tool and add it to the module's `*Tools` array
6. Add the new array to `src/tools/registry.ts` if creating a new category
7. Use `Result<T, McpError>` return types in underlying API modules
8. Add shared constants to `src/constants.ts` if needed

#### Adding New API Endpoints

1. Add endpoint URL to `src/utils/api-config.ts`
2. Create a function in the appropriate `src/api/*.ts` module
3. Use `httpRequest()` from `src/utils/http.ts` for automatic retry and timeout handling
4. Return `Result<T, McpError>` for type-safe error handling

#### Capturing New API Endpoints

Run `npm run research`, perform actions in Teams, and check the terminal output for captured requests.

## Troubleshooting

### Session/Token Expired
If API calls fail with authentication errors:
1. Call `teams_login` with `forceNew: true`
2. Or delete the config directory (`~/.msteams-mcp/` on macOS/Linux, `%APPDATA%\msteams-mcp\` on Windows) and run `npm run cli -- login`

### Browser Won't Launch (for login)
- Ensure you have Chrome (macOS/Linux) or Edge (Windows) installed
- On Windows, Edge should be pre-installed; try updating Windows if missing
- On macOS/Linux, install Chrome from https://www.google.com/chrome/
- Alternatively, download Playwright's bundled browser: `npx playwright install chromium`
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

## Reference

### File Locations

Session files are stored in a user-specific config directory to ensure consistency regardless of how the server is invoked (npx, global install, local dev, etc.):

- **macOS/Linux**: `~/.msteams-mcp/`
- **Windows**: `%APPDATA%\msteams-mcp\` (e.g., `C:\Users\name\AppData\Roaming\msteams-mcp\`)

Contents:
- `session-state.json` (encrypted browser session)
- `token-cache.json` (encrypted OAuth tokens)
- `.user-data/` (browser profile)

Legacy session files from the project root (`./session-state.json`) are automatically migrated to the new location on first read.

Development-only files (created in project root):
- **Debug output**: `./debug-output/` (gitignored, screenshots and HTML dumps)

Development files:
- **API reference**: `./docs/API-REFERENCE.md`

### API Endpoints

From research, Teams uses these primary APIs:

#### Search & Query
| Endpoint | Purpose |
|----------|---------|
| `substrate.office.com/searchservice/api/v2/query` | Full message search with pagination |
| `substrate.office.com/search/api/v1/suggestions` | People/message typeahead |
| `substrate.office.com/search/api/v1/suggestions?scenario=peoplecache` | Frequent contacts list |
| `substrate.office.com/search/api/v1/suggestions?domain=TeamsChannel` | Organisation-wide channel search |

#### Messages
| Endpoint | Purpose |
|----------|---------|
| `teams.microsoft.com/api/csa/{region}/api/v1/containers/{id}/posts` | Channel messages |
| `teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{id}/messages` | Send/receive messages |
| `teams.microsoft.com/api/chatsvc/{region}/v1/threads/{id}/annotations` | Reactions, read status |
| `teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders` | Favorites/pinned chats |
| `teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{id}/rcmetadata/{mid}` | Save/unsave messages |
| `teams.microsoft.com/api/chatsvc/{region}/v1/threads/{id}/consumptionhorizons` | Get read receipts |
| `teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{id}/properties?name=consumptionhorizon` | Mark as read |

#### People & Profile
| Endpoint | Purpose |
|----------|---------|
| `nam.loki.delve.office.com/api/v2/person` | Detailed person profile |
| `nam.loki.delve.office.com/api/v1/schedule` | Working hours, availability |
| `nam.loki.delve.office.com/api/v1/oofstatus` | Out of office status |
| `teams.microsoft.com/api/mt/part/{region}/beta/users/fetch` | Batch user lookup |

#### Files & Attachments
| Endpoint | Purpose |
|----------|---------|
| `substrate.office.com/AllFiles/api/users(...)/AllShared` | Files shared in conversation |

Regional identifiers: `amer`, `emea`, `apac`

See `docs/API-REFERENCE.md` for full endpoint documentation with request/response examples.

### Possible Tools

Based on API research, these tools could be implemented:

| Tool | API | Difficulty |
|------|-----|------------|
| `teams_get_person` | Delve person API | Easy |
| `teams_get_files` | AllFiles API | Medium |

**Known Limitations:**
- **Chat list** - Partially addressed by `teams_get_favorites` (pinned chats) and `teams_get_frequent_contacts` (common contacts), but no full chat list API
- **Presence/Status** - Real-time via WebSocket, not HTTP
- **Calendar** - Outlook APIs exist but need separate research

## Dependencies

- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `playwright`: Browser automation
- `zod`: Runtime input validation
- `vitest`: Unit testing framework (dev)
