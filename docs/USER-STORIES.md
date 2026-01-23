# Teams MCP User Stories

This document defines user stories and personas to guide development of the Teams MCP. Each story maps to specific API capabilities needed.

---

## Personas

### 🧑‍💼 Alex - Busy Manager
- Receives 100+ messages daily across multiple channels
- Needs to quickly catch up on what matters
- Often works across time zones, misses real-time conversations
- Wants AI to help prioritise and respond

### 👩‍💻 Sam - Developer
- Part of 10+ project channels
- Gets tagged in technical discussions
- Needs to find past decisions and context quickly
- Wants to automate routine responses

### 🧑‍🎨 Jordan - Creative Lead
- Collaborates with multiple teams
- Shares files and feedback frequently
- Needs to track project updates across channels
- Wants summaries rather than reading everything

---

## User Stories

### 1. Search & Reply

#### 1.1 Find and reply to a message
> "Find the message from Sarah about the budget review and reply saying I'll review it tomorrow."

**Flow:**
1. Search for messages matching "budget review from:sarah"
2. Display results with context
3. User confirms which message
4. Send reply to that conversation

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |
| `teams_get_thread_context` | ❌ Needed - get surrounding messages |

**Gap:** Need to get conversationId from search result to reply in correct thread.

---

#### 1.2 Search with date filters
> "Find messages from last week mentioning 'deployment'"

**Flow:**
1. Search with `sent:lastweek deployment`
2. Return matching messages

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented (operators work) |

**Status:** ✅ Works now with search operators

---

### 2. Catch Up & Prioritise

#### 2.1 Review questions asked of me
> "Show me any questions people have asked me today that I haven't answered."

**Flow:**
1. Search for messages mentioning me with question marks
2. Filter to unanswered (no reply from me after)
3. Prioritise by sender importance/urgency

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_get_me` | ❌ Needed - get my user ID/email |
| `teams_get_thread` | ❌ Needed - check if I replied |

**Gap:** Need user identity and thread reply detection.

---

#### 2.2 Catch up on unread messages
> "What unread messages do I have?"

**Flow:**
1. Get list of conversations with unread counts
2. Fetch unread messages from each
3. Summarise or list

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_unreads` | ❌ Needed |
| `teams_get_conversation_messages` | ❌ Needed |

**Gap:** Unread state is client-side in Teams. May need to track read position via `consumptionhorizon` API.

---

#### 2.3 Catch up on a channel
> "Summarise what happened in #project-alpha today"

**Flow:**
1. Find channel by name
2. Get recent messages from channel
3. Generate summary

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ❌ Needed |
| `teams_get_channel_posts` | ❌ Needed |

**Gap:** Channel discovery and message retrieval not yet implemented.

---

### 3. Favourites & Navigation

#### 3.1 List favourite channels
> "Show me my pinned/favourite channels"

**Flow:**
1. Get user's favourite channels list
2. Display with recent activity indicator

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_favorites` | ❌ Needed |

**Gap:** Favourites loaded at Teams startup, not via separate API. May need to capture from initial page load or localStorage.

---

#### 3.2 List recent chats
> "Who have I been chatting with recently?"

**Flow:**
1. Get recent 1:1 and group chats
2. Show with last message preview

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_recent_chats` | ❌ Needed |

**Gap:** Chat list loaded at startup. Similar challenge to favourites.

---

### 4. People & Profiles

#### 4.1 Find and message someone
> "Send a message to John Smith asking about the project status"

**Flow:**
1. Search for person by name
2. Get their conversation ID
3. Send message

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ❌ Needed |
| `teams_get_or_create_chat` | ❌ Needed |
| `teams_send_message` | ✅ Implemented |

**Gap:** Need people search and chat creation.

---

#### 4.2 Check someone's availability
> "Is Sarah available for a call right now?"

**Flow:**
1. Find person
2. Get their presence/availability status

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ❌ Needed |
| `teams_get_presence` | ❌ Needed (WebSocket-based) |

**Gap:** Presence is real-time via WebSocket, not HTTP API.

---

#### 4.3 Get my profile
> "What's my Teams email address?"

**Flow:**
1. Get current user profile

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_me` | ❌ Needed (Delve API available) |

**Status:** API discovered, implementation pending.

---

### 5. Notifications & Activity

#### 5.1 Review activity feed
> "Show me my recent notifications"

**Flow:**
1. Get activity/notification feed
2. Display with context

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_activity` | ❌ Needed |

**Gap:** Activity at `48:notifications` endpoint but format unclear.

---

#### 5.2 Find mentions of me
> "Show messages where I was @mentioned this week"

**Flow:**
1. Search for @mentions of current user
2. Filter by date

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_get_me` | ❌ Needed |

**Gap:** Need user identity to search for @mentions.

---

### 6. Files & Attachments

#### 6.1 Find shared files
> "Find the Excel file Sarah shared last week"

**Flow:**
1. Search for file by name/sender
2. Return download link or preview

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_files` | ❌ Needed |
| `teams_get_shared_files` | ❌ Needed (AllFiles API available) |

**Status:** API discovered, implementation pending.

---

### 7. Calendar & Meetings (Stretch Goal)

#### 7.1 Check upcoming meetings
> "What meetings do I have today?"

**Flow:**
1. Get calendar events for today
2. Include Teams meeting links

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_calendar` | ❌ Needed (Outlook API) |

**Gap:** Requires Outlook calendar APIs, separate auth scope.

---

## Implementation Priority

Based on user value and API readiness:

### Phase 1 - Quick Wins (APIs ready)
| Story | Tools Needed | Effort |
|-------|-------------|--------|
| 1.2 Search with filters | None (works now) | ✅ Done |
| 4.3 Get my profile | `teams_get_me` | Low |
| 4.1 Find person | `teams_search_people` | Low |
| 5.2 Find @mentions | `teams_get_me` | Low |

### Phase 2 - Core Functionality
| Story | Tools Needed | Effort |
|-------|-------------|--------|
| 1.1 Find & reply | Thread context | Medium |
| 2.3 Channel catchup | Channel posts | Medium |
| 6.1 Find files | Files API | Medium |

### Phase 3 - Advanced Features
| Story | Tools Needed | Effort |
|-------|-------------|--------|
| 2.1 Unanswered questions | Thread analysis | High |
| 2.2 Unread messages | Consumption horizon | High |
| 3.1 Favourites | Startup capture | High |

### Phase 4 - Stretch Goals
| Story | Tools Needed | Effort |
|-------|-------------|--------|
| 4.2 Presence | WebSocket | Very High |
| 7.1 Calendar | Outlook APIs | High |

---

## Next Steps

1. **Implement `teams_get_me`** - Unlocks @mention searches and reply context
2. **Implement `teams_search_people`** - Enables "message X person" flows
3. **Implement `teams_get_channel_posts`** - Enables channel catchup
4. **Add conversationId extraction** - Enable replying to search results

---

## Notes

### Search Operators (Already Working)
```
from:john.smith@company.com    # Messages from person
to:me                          # Messages sent to you
in:general                     # Messages in channel
sent:today                     # Messages from today
sent:lastweek                  # Messages from last week
hasattachment:true             # Messages with files
```

Combine operators: `from:sarah sent:lastweek hasattachment:true`

### Conversation IDs
- `48:notes` - Self-chat (notes to yourself)
- `48:notifications` - Activity feed
- `19:xxx@thread.tacv2` - Channel conversation
- `19:xxx@unq.gbl.spaces` - 1:1 or group chat
