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
| `in:` | `in:general` | Messages in a channel |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | By date (explicit dates only) |
| `subject:` | `subject:budget` | In message subject |
| `hasattachment:true` | - | Has files attached |
| `"Name"` | `"Smith, John"` | Find @mentions (name in quotes) |
| `NOT` | `NOT from:user@co.com` | Exclude results |

**Note:** `@me`, `from:me`, `to:me`, and `mentions:me` do NOT work. Use `teams_get_me` to get your actual email/display name, then use those values.

**⚠️ Date Operator Limitations:**
The `sent:` operator only works with explicit dates (e.g., `sent:2026-01-20` or `sent:>=2026-01-15`). Named shortcuts like `sent:lastweek`, `sent:today`, `sent:thisweek`, `sent:thismonth` do NOT work - they return 0 results. Results are sorted by recency, so date filters are often unnecessary.

**Finding Mentions:**
To find messages where you were @mentioned, search for your display name in quotes:
```
"Macdonald, Rob"              # Find mentions of you
"Macdonald, Rob" from:diego   # Mentions from Diego
```

**Combining Operators:**
```
from:john sent:>=2026-01-18   # John's messages since Jan 18
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

### 1.4 Channel Search (Organisation-wide) ✅ IMPLEMENTED

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel`

**Use Case:** Search for Teams channels across the entire organisation (not just channels the user is a member of).

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": {
        "QueryString": "testing",
        "DisplayQueryString": "testing"
      },
      "EntityType": "TeamsChannel",
      "Size": 10
    }
  ],
  "cvid": "uuid",
  "logicalId": "uuid"
}
```

**Response:**
```json
{
  "Groups": [
    {
      "Suggestions": [
        {
          "Name": "AI In Testing",
          "ThreadId": "19:ca554e7ce3ab4a2f8099765fba3079bf@thread.tacv2",
          "TeamName": "PE AI Excellence",
          "GroupId": "df865310-bf69-4f1b-8dc7-ebd0cbfa090f",
          "EntityType": "ChannelSuggestion",
          "ChannelType": "Standard",
          "Text": "testing",
          "PropertyHits": ["Name"],
          "ReferenceId": "uuid"
        }
      ]
    }
  ]
}
```

**Key Fields:**
- `Name`: Channel display name
- `ThreadId`: Conversation ID for use with messaging/thread APIs
- `TeamName`: Parent team's display name
- `GroupId`: Team's Azure AD group ID
- `ChannelType`: "Standard", "Private", or "Shared"

---

## 2. Channel & Chat APIs

### 2.1 Teams List (User's Joined Teams) ✅ NEW

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me`

**Query Parameters:**
- `isPrefetch=false`
- `enableMembershipSummary=true`
- `supportsAdditionalSystemGeneratedFolders=true`
- `supportsSliceItems=true`
- `enableEngageCommunities=false`

**Use Case:** Get all teams and channels the user is a member of. This is the main endpoint for discovering teams/channels.

**Authentication:** Requires both:
- `Authentication: skypetoken={skypeToken}` header
- `Authorization: Bearer {csaToken}` header (from MSAL with chatsvcagg.teams.microsoft.com audience)

**Response:**
```json
{
  "conversationFolders": {
    "folderHierarchyVersion": 1769200822270,
    "conversationFolders": [
      {
        "id": "folder-guid",
        "sortType": "UserDefinedCustomOrder",
        "name": "Folder Name",
        "folderType": "UserCreated",
        "conversationFolderItems": [
          {
            "conversationId": "19:channelId@thread.tacv2",
            "createdTime": 1753172521981,
            "lastUpdatedTime": 1753172521981
          }
        ]
      }
    ]
  },
  "teams": [
    {
      "threadId": "19:teamId@thread.tacv2",
      "displayName": "Team Name",
      "description": "Team description",
      "pictureETag": "etag-value",
      "isFavorite": false,
      "channels": [
        {
          "id": "19:channelId@thread.tacv2",
          "displayName": "General",
          "description": "Channel description",
          "isFavorite": true,
          "membershipType": "standard"
        }
      ]
    }
  ]
}
```

**Notes:**
- The `teams` array contains all teams with their channels nested inside
- Each team has a `threadId` (the team's root conversation ID) and a `channels` array
- Channel `id` can be used with other APIs to get posts, send messages, etc.
- The `conversationFolders` section contains user-created folders and pinned items

---

### 2.2 Channel Posts

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

### 2.3 Conversation Details

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

### 2.4 Thread Annotations (Reactions, Read Status)

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

### 2.5 Message Reactions (Emotions)

**Add Reaction Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Remove Reaction Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Use Case:** Add or remove emoji reactions from messages.

**Add Reaction Request:**
```json
{
  "emotions": {
    "key": "like",
    "value": 1769429691997
  }
}
```

**Remove Reaction Request:**
```json
{
  "emotions": {
    "key": "like"
  }
}
```

**Key Format:**
- Standard Teams emojis: just the name (e.g., `like`, `heart`, `laugh`, `surprised`, `sad`, `angry`, `elephant`)
- Custom/animated org emojis: `{name};{storage-id}` (e.g., `octi-search;0-wus-d10-66fac2a3b0cda332435c21a14485efe7`)

**Common Reaction Keys:**
| Emoji | Key |
|-------|-----|
| 👍 | `like` |
| ❤️ | `heart` |
| 😂 | `laugh` |
| 😮 | `surprised` |
| 😢 | `sad` |
| 😠 | `angry` |

**Notes:**
- The `value` field is a timestamp in milliseconds
- Standard emoji search is client-side - catalog bundled in JS
- Authentication uses the same `skypetoken_asm` cookie as messaging

---

### 2.6 Custom Emoji Metadata

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/customemoji/metadata`

**Use Case:** Get the list of custom organisation emojis (not standard Teams emojis).

**Response:**
```json
{
  "continuationToken": "1769429894",
  "categories": [
    {
      "id": "customEmoji",
      "title": "Custom Emoji",
      "description": "Custom Emoji",
      "emoticons": [
        {
          "id": "angrysteam;0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "documentId": "0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "shortcuts": ["angrysteam"],
          "description": "angrysteam",
          "createdOn": 1721360079509,
          "isDeleted": false
        }
      ]
    }
  ]
}
```

**Notes:**
- The `id` field is what you use as the reaction key (includes the storage reference)
- `shortcuts` are the search terms that match this emoji
- Standard Teams emojis are bundled in JS, not available via API
- Image URL pattern: `https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/{emoji-id}/default/20_f.png`

---

### 2.7 Consumption Horizons (Read Receipts)

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

### 2.8 Conversation Folders (Favorites/Pinned)

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

### 2.9 Saved Messages (Bookmarks)

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

### 4.2 Reply to Thread (Channel)

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{channelId};messageid={threadRootMessageId}/messages`

**Use Case:** Reply to an existing thread in a channel. The `;messageid=` suffix in the URL path indicates this is a reply to a specific thread.

**URL Pattern Differences:**

| Action | URL Path |
|--------|----------|
| New channel post | `conversations/{channelId}/messages` |
| Reply to thread | `conversations/{channelId};messageid={threadRootId}/messages` |
| Chat message | `conversations/{chatId}/messages` |

**Request:** Same structure as regular message send.

**Notes:**
- The `threadRootMessageId` is the timestamp/ID of the first message in the thread
- The `;messageid=` is URL-encoded as `%3Bmessageid%3D` in the actual request
- The `conversationLink` in the body should also include `;messageid={id}` for thread replies
- Chats (1:1, group, meeting) don't use threading - all messages go to the flat conversation

---

### 4.3 Typing Indicator

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

### 4.4 Edit Message ✅ IMPLEMENTED

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}`

**Use Case:** Edit an existing message. You can only edit your own messages.

**Request:**
```json
{
  "id": "{messageId}",
  "type": "Message",
  "conversationid": "{conversationId}",
  "content": "<p>Updated message content</p>",
  "messagetype": "RichText/Html",
  "contenttype": "text",
  "imdisplayname": "Display Name"
}
```

**Response:** `200 OK` (empty or minimal body)

**Notes:**
- Edited messages have a `skypeeditedid` field when fetched
- The API returns 403 Forbidden if you try to edit someone else's message

---

### 4.5 Delete Message ✅ IMPLEMENTED

**Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}?behavior=softDelete`

**Use Case:** Delete a message (soft delete). You can only delete your own messages, unless you are a channel owner/moderator.

**Response:** `200 OK` with `null` body

**Notes:**
- This is a soft delete - the message is flagged, not removed from the database
- Search API filters deleted messages with `AND NOT (isClientSoftDeleted:TRUE)`
- Channel owners/moderators can delete other users' messages
- The API returns 403 Forbidden for unauthorised delete attempts

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

### 7.1 Activity Feed (Notification Messages) ✅ IMPLEMENTED

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Anotifications/messages`

**Query Parameters:**
- `view=msnp24Equivalent` - Response format (same as other message endpoints)
- `pageSize=50` - Number of activity items to return (max ~200)
- `syncState={base64State}` - For incremental sync (optional)

**Use Case:** Get the user's activity feed - mentions, reactions, replies, announcements, etc.

**Authentication:** Same as other chatsvc endpoints - requires `skypetoken` and Bearer token.

**Request Headers:**
```
Authentication: skypetoken={skypeToken}
Authorization: Bearer {authToken}
Content-Type: application/json
```

**Response:**
```json
{
  "messages": [
    {
      "id": "1769276832046",
      "originalarrivaltime": "2026-01-24T18:47:12.046Z",
      "composetime": "2026-01-24T18:47:12.046Z",
      "messagetype": "RichText/Html",
      "contenttype": "text",
      "content": "<p>Activity content here</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John",
      "conversationid": "19:meeting_abc123@thread.v2",
      "threadtopic": "Weekly Standup",
      "clientmessageid": "12345678901234567890"
    }
  ],
  "syncState": "base64EncodedState..."
}
```

**Activity Types (identified via `messagetype` and content patterns):**
| Type | Identification |
|------|----------------|
| @Mention | Content contains `<span itemtype="http://schema.skype.com/Mention">` |
| Reaction | `messagetype` contains reaction identifier |
| Reply | Standard message in a thread context |
| Announcement | Channel announcement formatting |

**Notes:**
- The conversation ID `48:notifications` is a special feed (like `48:notes` for self-chat)
- Activity items include context like `conversationid` to navigate to the source
- The `threadtopic` field contains the conversation/meeting name when available
- Use `syncState` from response for efficient incremental polling

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

Based on discovered APIs, here are the current tool implementation status:

### ✅ Implemented

| Tool | API | Notes |
|------|-----|-------|
| `teams_search` | Substrate v2 query | Full-text search with pagination |
| `teams_login` | Browser automation | Manual login flow |
| `teams_status` | Token check | Auth status for all APIs |
| `teams_send_message` | chatsvc messages API | Send to any conversation |
| `teams_reply_to_thread` | chatsvc messages API | Reply to channel threads |
| `teams_edit_message` | chatsvc messages API | Edit own messages |
| `teams_delete_message` | chatsvc messages API | Soft delete own messages |
| `teams_get_me` | JWT token extraction | Current user profile |
| `teams_get_favorites` | conversationFolders API | Pinned/favourite conversations |
| `teams_add_favorite` | conversationFolders API | Pin a conversation |
| `teams_remove_favorite` | conversationFolders API | Unpin a conversation |
| `teams_save_message` | rcmetadata API | Bookmark a message |
| `teams_unsave_message` | rcmetadata API | Remove bookmark |
| `teams_search_people` | Substrate suggestions | Find people by name/email |
| `teams_get_frequent_contacts` | Substrate peoplecache | Frequently contacted people |
| `teams_get_thread` | chatsvc messages API | Messages from any conversation |
| `teams_find_channel` | Teams List + Substrate | Hybrid channel search |
| `teams_get_chat` | Computed from user IDs | 1:1 conversation ID (no API call) |
| `teams_get_unread` | chatsvc consumptionhorizons | Unread counts (aggregate or per-conversation) |
| `teams_mark_read` | chatsvc consumptionhorizon | Mark conversation read up to message |
| `teams_get_activity` | chatsvc 48:notifications | Activity feed (mentions, reactions, replies) |

### 🔜 Ready to Implement

| Tool | API | Notes |
|------|-----|-------|
| `teams_get_person` | Delve person API | Get specific person's details |
| `teams_get_files` | AllFiles API | List files shared in a conversation |

### ⚠️ Needs More Research

| Tool | Notes |
|------|-------|
| `teams_list_chats` | No dedicated API found - may need initial load capture |
| `teams_get_saved_messages` | No single endpoint - saved flag is per-message in rcMetadata |
| `teams_calendar` | Outlook calendar APIs available - need to extract meetings |

### 📝 Implementation Notes

**Channel Posts via `teams_get_thread`:**

The `teams_get_channel_posts` tool is no longer needed as a separate implementation. The existing `teams_get_thread` tool works with channel IDs:

1. Use `teams_find_channel` to discover channels by name → returns `channelId` (format: `19:xxx@thread.tacv2`)
2. Use `teams_get_thread` with that `channelId` to get messages

The CSA containers API (section 2.2) provides additional features like `threadedPostsOnly` for top-level posts only, but for most use cases `teams_get_thread` via chatsvc is sufficient.

**1:1 Chat Conversation IDs (Discovered):**

Research revealed that 1:1 chat conversation IDs are **predictable** and don't require a creation API:

**Format:** `19:{userId1}_{userId2}@unq.gbl.spaces`

Where:
- `userId1` and `userId2` are Azure AD object IDs (GUIDs)
- The IDs are **sorted lexicographically** (ensures both participants get the same ID)
- The conversation is created implicitly when the first message is sent

**Example:**
- Your ID: `ab76f827-27e2-4c67-a765-f1a53145fa24`
- Other person: `5817f485-f870-46eb-bbc4-de216babac62`
- Since '5' < 'a' alphabetically: `19:5817f485-..._ab76f827-...@unq.gbl.spaces`

**Implementation:** The `teams_get_chat` tool computes this ID from any user identifier format (MRI, ID with tenant, or raw GUID). Use the result with `teams_send_message`.

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
