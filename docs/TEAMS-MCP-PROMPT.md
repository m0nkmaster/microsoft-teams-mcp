# Teams MCP Usage Prompt

Use this prompt when you want an AI assistant to interact with Microsoft Teams on your behalf.

> **Manual Testing**: This document also serves as a test script before releasing new versions. Ask an AI agent to "run this script to test the MCP" and it will exercise all documented tools, verifying authentication, search, messaging, and management features work correctly.

---

## Available Tools

### Authentication & Status
| Tool | Purpose |
|------|---------|
| `teams_status` | Check authentication status and token expiry |
| `teams_get_me` | Get your profile (email, name, ID) |
| `teams_login` | Trigger manual login if session expired |

### Search & Discovery
| Tool | Purpose |
|------|---------|
| `teams_search` | Search messages with query operators |
| `teams_find_channel` | Find channels by name (your teams + org-wide) |
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (for name resolution) |

### Reading Messages
| Tool | Purpose |
|------|---------|
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_get_unread` | Get unread status (aggregate or specific conversation) |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies) |

### Sending Messages
| Tool | Purpose |
|------|---------|
| `teams_send_message` | Send a message (defaults to self-notes) |
| `teams_reply_to_thread` | Reply to a channel message as threaded reply |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |

### Message Management
| Tool | Purpose |
|------|---------|
| `teams_edit_message` | Edit your own messages |
| `teams_delete_message` | Delete your own messages (soft delete) |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark |
| `teams_mark_read` | Mark conversation as read up to a message |

### Favourites
| Tool | Purpose |
|------|---------|
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |

---

## Search Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:user@company.com` | Messages from a person |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | Messages by date (explicit dates only) |
| `in:` | `in:channel-name` | Messages in a channel |
| `"Name"` | `"John Smith"` | Find @mentions |
| `NOT` | `NOT from:user@company.com` | Exclude results |
| `hasattachment:` | `hasattachment:true` | Messages with files |

**Important**: `@me`, `from:me`, `to:me` do NOT work. Use `teams_get_me` first to get your actual email/name. Also `sent:lastweek`, `sent:today`, `sent:thisweek` do NOT work - use explicit dates (e.g., `sent:>=2026-01-18`) or omit since results are sorted by recency.

---

## Common Workflows

### Find messages mentioning me
```
1. teams_get_me → get displayName and email
2. teams_search "Display Name" NOT from:my.email@company.com
```

### Find and message someone
```
1. teams_search_people "person name" → get their user ID
2. teams_get_chat userId → get conversation ID
3. teams_send_message content="Hello" conversationId="..."
```

### Reply to a channel thread
```
1. teams_search or teams_get_thread → get conversationId and messageId
2. teams_reply_to_thread content="Reply" conversationId="..." messageId="..."
```

### Test reply and delete (for manual testing)
```
1. teams_get_me → get your email
2. teams_search from:your.email@company.com → find your own channel message
3. teams_reply_to_thread content="Test reply" → reply to your own message
4. teams_delete_message → delete the test reply to clean up
```

### Check unread messages
```
1. teams_get_unread → aggregate unread across favourites
2. teams_get_thread conversationId="..." → read the messages
3. teams_mark_read conversationId="..." messageId="..." → mark as read
```

### Find a channel to read
```
1. teams_find_channel "channel name" → get channelId
2. teams_get_thread conversationId="channelId" → read messages
```

---

## Known Limitations

| Limitation | Details |
|------------|---------|
| Self-notes (`48:notes`) | Cannot edit/delete messages in self-chat |
| Save message | Only works on root messages, not thread replies |
| Unread on channels | May fail ACL check; works reliably for chats/meetings |
| Token expiry | Tokens last ~1 hour; call `teams_login` to refresh |

---

## Safety Guidelines

- **Never send messages to others** without explicit user confirmation
- **Default to self-notes** (`48:notes`) for testing or drafts
- **Verify recipients** before sending by confirming email/name
- **Be cautious with delete** - it's a soft delete but still removes content

---

## Example Prompts

### Catch up on Teams
> "Check my Teams for any unread messages or mentions. Summarise what needs my attention."

### Find information
> "Search Teams for recent discussions about [topic]. Who's been involved and what are the key points?"

### Draft a message
> "Help me draft a message to [person] about [topic]. Save it to my notes first so I can review."

### Channel monitoring
> "Check the [channel name] channel for recent activity and summarise any important updates."

### People lookup
> "Find [person name]'s contact details and check if I have any recent conversations with them."
