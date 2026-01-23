# Teams Web App API Research

This document captures findings from researching the Microsoft Teams web application.

## Table of Contents

1. [Search APIs](#1-search-apis)
2. [Channel & Chat APIs](#2-channel--chat-apis)
3. [People APIs](#3-people-apis)
4. [Messaging APIs](#4-messaging-apis)
5. [Calendar & Scheduling](#5-calendar--scheduling-apis)
6. [Files & Attachments](#6-files--attachments-apis)
7. [Notifications & Activity](#7-notifications--activity-apis)
8. [Authentication](#authentication)
9. [Regional Variations](#regional-variations)
10. [Potential MCP Tools](#potential-mcp-tools)
11. [Not Yet Captured](#not-yet-captured)

---

## 1. Search APIs

### 1.1 Substrate v2 Query (Full Search with Pagination) ✅ IMPLEMENTED

**Endpoint:** `POST https://substrate.office.com/searchservice/api/v2/query`

**Use Case:** Full-text search of Teams messages with pagination support.

**Request:**
```json
{
  "entityRequests": [
    {
      "entityType": "Message",
      "contentSources": ["Teams"],
      "fields": [
        "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String",
        "Extension_SkypeSpaces_ConversationPost_Extension_FileData_String",
        "Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String"
      ],
      "propertySet": "Optimized",
      "query": {
        "queryString": "search term AND NOT (isClientSoftDeleted:TRUE)",
        "displayQueryString": "search term"
      },
      "from": 0,
      "size": 25,
      "topResultsCount": 5
    }
  ],
  "QueryAlterationOptions": {
    "EnableAlteration": true,
    "EnableSuggestion": true,
    "SupportedRecourseDisplayTypes": ["Suggestion"]
  },
  "cvid": "uuid",
  "logicalId": "uuid",
  "scenario": {
    "Dimensions": [
      {"DimensionName": "QueryType", "DimensionValue": "Messages"},
      {"DimensionName": "FormFactor", "DimensionValue": "general.web.reactSearch"}
    ],
    "Name": "powerbar"
  }
}
```

**Response:**
```json
{
  "EntitySets": [
    {
      "ResultSets": [
        {
          "Total": 4307,
          "Results": [
            {
              "Id": "AAMkA...",
              "ReferenceId": "uuid.1000.1",
              "HitHighlightedSummary": "Message with <c0>highlights</c0>...",
              "Source": {
                "Summary": "Plain text content",
                "From": {
                  "EmailAddress": {
                    "Name": "Smith, John",
                    "Address": "john.smith@company.com"
                  }
                }
              }
            }
          ]
        }
      ]
    }
  ]
}
```

**Pagination:**
- `from`: Starting offset (0, 25, 50, ...)
- `size`: Page size (default 25, max ~50)
- Response includes `Total` count for accurate pagination

**Search Operators (passed through to queryString):**
| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:john.smith@company.com` | Messages from a person |
| `to:` | `to:me` | Messages sent to you |
| `in:` | `in:general` | Messages in a channel |
| `sent:` | `sent:today`, `sent:lastweek` | By date |
| `subject:` | `subject:budget` | In message subject |
| `hasattachment:true` | - | Has files attached |

**Finding Mentions:**
To find messages where you were @mentioned, search for your display name in quotes:
```
"Macdonald, Rob"              # Find mentions of you
"Macdonald, Rob" from:diego   # Mentions from Diego
```

**Combining Operators:**
```
from:john sent:lastweek # John's messages last week
```

---

### 1.2 Substrate v1 Suggestions (Autocomplete/Type-ahead)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar`

**Use Case:** Real-time search suggestions as user types.

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": {
        "QueryString": "rob",
        "DisplayQueryString": "rob"
      },
      "EntityType": "People",
      "Size": 5,
      "Fields": ["Id", "MRI", "DisplayName", "EmailAddresses", "JobTitle", "Department"]
    }
  ]
}
```

**Response:**
```json
{
  "Groups": [
    {
      "Suggestions": [
        {
          "Id": "uuid@tenant",
          "DisplayName": "Smith, John",
          "GivenName": "John",
          "Surname": "Smith",
          "EmailAddresses": ["user@company.com"],
          "CompanyName": "Company Name",
          "Department": "Engineering",
          "JobTitle": "Senior Engineer"
        }
      ]
    }
  ]
}
```

---

### 1.3 People Cache (Frequent Contacts)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache`

**Use Case:** Get list of frequently contacted people (can serve as "favorites" proxy).

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": { "QueryString": "", "DisplayQueryString": "" },
      "EntityType": "People",
      "Size": 500,
      "Fields": ["Id", "MRI", "DisplayName", "EmailAddresses", "GivenName", "Surname", "CompanyName", "JobTitle"]
    }
  ]
}
```

**Response:** Same format as suggestions, returns ranked list of contacts.

---

## 2. Channel & Chat APIs

### 2.1 Channel Posts

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/containers/{containerId}/posts`

**Query Parameters:**
- `threadedPostsOnly=true` - Only top-level posts
- `pageSize=20` - Number of posts
- `teamId={teamId}` - Parent team ID
- `includeRcMetadata=true` - Include read/consumption metadata

**Response:**
```json
{
  "posts": [
    {
      "containerId": "19:channelId@thread.tacv2",
      "id": "1769189921704",
      "latestMessageTime": "2026-01-23T17:54:43.263Z",
      "message": {
        "messageType": "RichText/Html",
        "content": "<p>Message content with <span itemtype=\"http://schema.skype.com/Mention\">@mentions</span></p>",
        "fromFamilyNameInToken": "Smith",
        "fromGivenNameInToken": "John"
      }
    }
  ]
}
```

---

### 2.2 Conversation Details

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}?view=msnp24Equivalent`

**Use Case:** Get conversation metadata for a specific chat or channel.

**Response Structure:**
```json
{
  "id": "19:abc@thread.tacv2",
  "threadProperties": {
    "threadType": "topic",
    "productThreadType": "TeamsStandardChannel",
    "groupId": "guid",
    "topic": "Channel topic",
    "topicThreadTopic": "Channel Name",
    "spaceThreadTopic": "Team Name",
    "spaceId": "19:teamroot@thread.tacv2"
  },
  "members": [...],
  "lastMessage": {...}
}
```

**Conversation Type Identification:**

| Type | `threadType` | `productThreadType` | Name Source |
|------|--------------|---------------------|-------------|
| Standard Channel | `topic` | `TeamsStandardChannel` | `topicThreadTopic` |
| Team (General/Root) | `space` | `TeamsTeam` | `spaceThreadTopic` |
| Private Channel | `space` | `TeamsPrivateChannel` | `topicThreadTopic` |
| Meeting Chat | `meeting` | `Meeting` | `topic` |
| Group Chat | `chat` | `Chat` | `topic` or members |
| 1:1 Chat | `chat` | `OneOnOne` | Other participant |

**Key Fields:**
- `groupId`: Present for all Team-related conversations (channels, team root)
- `topicThreadTopic`: The channel name within a team
- `spaceThreadTopic`: The parent team's name
- `topic`: User-set topic for chats, or meeting title

---

### 2.3 Thread Annotations (Reactions, Read Status)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/annotations?messageIds={id1},{id2}`

**Use Case:** Get reactions and read status for specific messages.

**Response:**
```json
{
  "annotations": {
    "1769013008614": {
      "annotations": [
        {
          "mri": "8:orgid:user-guid",
          "time": 1769076365,
          "annotationType": "l2ch",
          "annotationGroup": "userMetaData"
        }
      ]
    }
  }
}
```

---

### 2.4 Consumption Horizons (Read Receipts)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/consumptionhorizons`

**Use Case:** Track what has been read in a conversation.

**Response:**
```json
{
  "id": "19:meeting_abc@thread.v2",
  "version": "1769191217184",
  "consumptionhorizons": []
}
```

---

### 2.5 Conversation Folders (Favorites/Pinned)

**Endpoint:** `POST https://teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders`

**Query Parameters:** `supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`

**Use Case:** Get pinned/favourite conversations, or add/remove items from folders.

**Request (Get all folders):**
```json
{}
```

**Request (Add to Favorites):**
```json
{
  "folderHierarchyVersion": 1769191147787,
  "actions": [
    {
      "action": "AddItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "19:conversationId@thread.v2"
    }
  ]
}
```

**Request (Remove from Favorites):**
```json
{
  "actions": [
    {
      "action": "RemoveItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "19:conversationId@thread.v2"
    }
  ]
}
```

**Response:**
```json
{
  "folderHierarchyVersion": 1769200822270,
  "conversationFolders": [
    {
      "id": "{tenantId}~{userId}~Favorites",
      "sortType": "UserDefinedCustomOrder",
      "name": "Favorites",
      "folderType": "Favorites",
      "conversationFolderItems": [
        {
          "conversationId": "19:abc@thread.tacv2",
          "createdTime": 1750768187119,
          "lastUpdatedTime": 1750768187119
        }
      ]
    }
  ]
}
```

**Notes:**
- The `folderHierarchyVersion` should be included from a previous response for updates
- Folder ID format: `{tenantId}~{userId}~{folderType}`
- Known folder types: `Favorites`

---

### 2.6 Saved Messages (Bookmarks)

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/rcmetadata/{messageId}`

**Use Case:** Save (bookmark) or unsave a message.

**Request (Save message):**
```json
{
  "s": 1,
  "mid": 1769200192761
}
```

**Request (Unsave message):**
```json
{
  "s": 0,
  "mid": 1769200192761
}
```

**Response:**
```json
{
  "conversationId": "19:f205342987334b6d9722c1d52b526400@thread.v2",
  "rootMessageId": 1769200192761,
  "rcMetadata": {
    "lu": 1769200800298,
    "s": 1
  }
}
```

**Notes:**
- `s`: Saved flag (1 = saved, 0 = not saved)
- `mid`: Message ID
- `lu`: Last updated timestamp
- To retrieve saved status for messages, use the posts API with `includeRcMetadata=true`
- There is no single "get all saved messages" endpoint; saved messages are tracked per-conversation

---

## 3. People APIs

### 3.1 Person Profile (Delve/Loki)

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v2/person`

**Query Parameters:** `smtp={email}&personaType=User&locale=en-gb`

**Use Case:** Get detailed profile for a person.

**Response:**
```json
{
  "person": {
    "names": [
      {
        "value": {
"displayName": "Smith, John",
        "givenName": "John",
        "surname": "Smith"
        },
        "source": "Organisation"
      }
    ],
    "emailAddresses": [
      {
        "value": {
          "name": "user@company.com",
          "address": "user@company.com"
        }
      }
    ]
  }
}
```

---

### 3.2 Profile Picture

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/users/{userId}/profilepicturev2/{mri}?size=HR96x96`

**Use Case:** Get user's profile picture in various sizes (HR64x64, HR96x96, HR196x196).

---

### 3.3 User Lookup (Batch)

**Endpoint:** `POST https://teams.microsoft.com/api/mt/part/{region}/beta/users/fetch`

**Request Body:**
```json
["8:orgid:{userId1}", "8:orgid:{userId2}"]
```

**Response:**
```json
{
  "value": [
    {
      "alias": "USERNAME",
      "mail": "user@domain.com",
      "displayName": "Display Name",
      "objectType": "User"
    }
  ]
}
```

---

### 3.4 Graph API User Query

**Endpoint:** `GET https://graph.microsoft.com/v1.0/users?$select=id,mail&$filter=mail+in("email@domain.com")`

**Use Case:** Look up user by email address.

**Response:**
```json
{
  "value": [
    {
      "id": "00000000-0000-0000-0000-000000000000",
      "mail": "user@company.com"
    }
  ]
}
```

---

## 4. Messaging APIs

### 4.1 Send Message

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages`

**Use Case:** Send a new message to a chat or channel.

**Request:**
```json
{
  "id": "-1",
  "type": "Message",
  "conversationid": "{conversationId}",
  "conversationLink": "blah/{conversationId}",
  "from": "8:orgid:{userId}",
  "fromUserId": "8:orgid:{userId}",
  "composetime": "2026-01-23T18:03:50.335Z",
  "originalarrivaltime": "2026-01-23T18:03:50.335Z",
  "content": "<p>Message content here</p>",
  "messagetype": "RichText/Html",
  "contenttype": "Text",
  "imdisplayname": "Display Name",
  "clientmessageid": "{uniqueId}"
}
```

**Response:**
```json
{
  "OriginalArrivalTime": 1769191432285
}
```

**Special Conversation IDs:**
- `48:notes` - Personal notes/self-chat
- `48:notifications` - Notifications feed

---

### 4.2 Typing Indicator

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages`

**Request:**
```json
{
  "content": "",
  "contenttype": "Application/Message",
  "messagetype": "Control/Typing"
}
```

---

## 5. Calendar & Scheduling APIs

### 5.1 Schedule / Next Availability

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/schedule`

**Query Parameters:** `smtp={email}&personaType=User`

**Response:**
```json
{
  "nextAvailability": {
    "utcDateTime": "0001-01-01T00:00:00",
    "currentStatus": "Free",
    "nextStatus": "Free"
  },
  "workingHoursCalendar": {
    "daysOfWeek": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "startTime": "08:00:00",
    "endTime": "17:00:00",
    "timeZone": {
      "name": "Greenwich Mean Time"
    }
  }
}
```

---

### 5.2 Out of Office Status

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/oofstatus`

**Query Parameters:** `smtp={email}&personaType=User`

**Response:**
```json
{
  "outOfOfficeState": "Disabled",
  "externalAudience": "Unknown",
  "emailAddress": "USER@COMPANY.COM"
}
```

---

### 5.3 Meeting Tabs

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/chats/{meetingThreadId}/tabs`

**Use Case:** Get tabs attached to a meeting chat (Polls, Notes, etc.)

**Response:**
```json
{
  "value": [
    {
      "id": "uuid",
      "name": "Polls",
      "appId": "uuid",
      "configuration": {
        "entityId": "TeamsMeetingPollPage",
        "contentUrl": "https://forms.office.com/..."
      }
    }
  ]
}
```

---

## 6. Files & Attachments APIs

### 6.1 Files Shared in Thread

**Endpoint:** `GET https://substrate.office.com/AllFiles/api/users('OID:{userId}@{tenantId}')/AllShared`

**Query Parameters:**
- `ThreadId={conversationId}`
- `ItemTypes=File&ItemTypes=Link`
- `PageSize=25`

**Response:**
```json
{
  "Items": [
    {
      "ItemType": "File",
      "FileData": {
        "FileName": "Meeting Recording.mp4",
        "FileExtension": "mp4",
        "WebUrl": "https://sharepoint.com/..."
      },
      "SharedByDisplayName": "Smith, John",
      "SharedBySmtp": "john.smith@company.com"
    },
    {
      "ItemType": "Link",
      "WeblinkData": {
        "WebUrl": "https://jira.company.com/...",
        "Title": "JIRA Ticket"
      }
    }
  ]
}
```

---

## 7. Notifications & Activity APIs

### 7.1 Notification Messages

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Anotifications/messages`

**Query Parameters:**
- `view=msnp24Equivalent|supportsMessageProperties`
- `pageSize=200`
- `syncState={base64State}`

**Use Case:** Get activity/notification feed messages.

---

### 7.2 Notification Settings

**Endpoint:** `GET https://teams.microsoft.com/api/nss/{region}/v1/me/notificationSettings/team/{teamId}/channel/{channelId}`

**Use Case:** Get notification preferences for a specific channel.

---

### 7.3 Update Read Status

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/properties?name=consumptionhorizon`

**Request:**
```json
{
  "consumptionhorizon": "{timestamp1};{timestamp2};{id}"
}
```

---

## Authentication

### Token Storage

Tokens are stored in browser localStorage under MSAL keys:
- Look for keys containing `SubstrateSearch` scope
- Tokens typically expire after ~1 hour

### Required Headers for Direct API Calls

```
Authorization: Bearer {token}
Origin: https://teams.microsoft.com
Referer: https://teams.microsoft.com/
Content-Type: application/json
```

### Session Persistence

The MCP server saves session state to `session-state.json` including:
- Cookies
- localStorage (contains MSAL tokens)
- sessionStorage

---

## Regional Variations

API URLs include regional identifiers:

| Region | Code | Example |
|--------|------|---------|
| Americas | `amer` | `/api/csa/amer/api/v1/...` |
| Europe/Middle East/Africa | `emea` | `/api/csa/emea/api/v1/...` |
| Asia Pacific | `apac` | `/api/csa/apac/api/v1/...` |

Also appears in:
- `/api/chatsvc/{region}/v1/...`
- `/api/mt/part/{region}-02/beta/...`
- `nam.loki.delve.office.com` (North America)

---

## Potential MCP Tools

Based on discovered APIs, here are potential tools to implement:

### ✅ Implemented

| Tool | API | Status |
|------|-----|--------|
| `teams_search` | Substrate v2 query | ✅ Done |
| `teams_login` | Browser automation | ✅ Done |
| `teams_status` | Token check | ✅ Done |
| `teams_send_message` | chatsvc messages API | ✅ Done (uses skypetoken cookies) |
| `teams_get_me` | JWT token extraction | ✅ Done |

### 🔜 Ready to Implement

| Tool | API | Notes |
|------|-----|-------|
| `teams_get_favorites` | conversationFolders API | Get pinned/favourite conversations |
| `teams_add_favorite` | conversationFolders API | Pin a conversation |
| `teams_remove_favorite` | conversationFolders API | Unpin a conversation |
| `teams_save_message` | rcmetadata API | Bookmark a message |
| `teams_unsave_message` | rcmetadata API | Remove bookmark from message |
| `teams_search_people` | Substrate suggestions | Find people by name/email |
| `teams_get_person` | Delve person API | Get specific person's details |
| `teams_get_channel_posts` | CSA containers API | Get messages from a channel |
| `teams_get_files` | AllFiles API | List files shared in a conversation |

### ⚠️ Needs More Research

| Tool | Notes |
|------|-------|
| `teams_list_chats` | No dedicated API found - may need initial load capture |
| `teams_get_saved_messages` | No single endpoint - saved flag is per-message in rcMetadata |
| `teams_activity` | 48:notifications exists but format unclear |
| `teams_calendar` | Outlook calendar APIs available - need to extract meetings |

### 🔮 Future Possibilities

| Tool | API | Notes |
|------|-----|-------|
| `teams_get_presence` | Not captured | Would need WebSocket/SignalR |
| `teams_unread` | Client-side filter | Use consumptionhorizon comparison |

---

## Not Yet Captured

These features likely need additional research:

1. **Favorites/Pinned Channels**
   - Loaded at Teams startup
   - May be in localStorage or initial boot response
   - Need fresh session research

2. **Full Chat List**
   - No `/chats/list` endpoint found
   - Teams pre-loads conversations
   - May use WebSocket for real-time updates

3. **Calendar Events**
   - Outlook APIs available (`outlook.office.com`)
   - Need to capture calendar view load

4. **Presence/Status**
   - Real-time via WebSocket
   - Not captured in HTTP intercept

5. **Teams/Channels List**
   - Structure loaded at startup
   - Need initial boot capture

---

## How to Conduct Research

1. Clear session: `rm session-state.json`
2. Run: `npm run research`
3. Log in to Teams when prompted
4. Perform actions (click Activity, Calendar, etc.)
5. Press Ctrl+C to stop
6. Check terminal output for captured requests/responses

For initial boot capture:
1. Clear all Teams data
2. Start research script
3. Complete full login flow
4. Watch for large initial payloads
