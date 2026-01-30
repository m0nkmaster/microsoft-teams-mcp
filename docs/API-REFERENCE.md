# Microsoft Teams Web API Reference

Complete reference for the undocumented Microsoft Teams APIs used by this MCP server. These are internal APIs discovered through browser network inspection, not the official Microsoft Graph API.

## Table of Contents

1. [Authentication](#authentication)
2. [Regional Variations](#regional-variations)
3. [Search APIs](#search-apis)
4. [People APIs](#people-apis)
5. [Virtual Conversations](#virtual-conversations)
6. [Messaging APIs](#messaging-apis)
7. [Conversation APIs](#conversation-apis)
8. [Activity & Notifications](#activity--notifications)
9. [Reactions & Emoji](#reactions--emoji)
10. [Calendar & Scheduling](#calendar--scheduling)
11. [Files & Attachments](#files--attachments)
12. [Common Gotchas](#common-gotchas)

---

## Authentication

Teams uses multiple authentication mechanisms depending on the API surface:

| Auth Type | Header | Source | Used By |
|-----------|--------|--------|---------|
| **Bearer (Substrate)** | `Authorization: Bearer {token}` | MSAL localStorage, `SubstrateSearch-Internal.ReadWrite` scope | Search, People |
| **Bearer (CSA)** | `Authorization: Bearer {csaToken}` | MSAL, `chatsvcagg.teams.microsoft.com` audience | Teams list, Favorites |
| **Skype Token** | `Authentication: skypetoken={token}` | Cookie `skypetoken_asm` | Messaging, Threads |

### Required Headers

Most endpoints require these headers:

```
Origin: https://teams.microsoft.com
Referer: https://teams.microsoft.com/
Content-Type: application/json
```

### Token Storage

Tokens are stored in browser localStorage under MSAL keys. Look for keys containing `SubstrateSearch` scope. Tokens typically expire after ~1 hour. MSAL only refreshes tokens when an API call requires them, not on page load.

### Session Persistence

Session state includes:
- Cookies (for messaging APIs)
- localStorage (contains MSAL tokens)
- sessionStorage

---

## Regional Variations

API URLs include regional identifiers based on the user's tenant location:

| Region | Code |
|--------|------|
| Americas | `amer` |
| Europe/Middle East/Africa | `emea` |
| Asia Pacific | `apac` |

**Usage patterns:**
- `/api/csa/{region}/api/v1/...`
- `/api/chatsvc/{region}/v1/...`
- `/api/mt/part/{region}/beta/...`
- `nam.loki.delve.office.com` (North America Delve APIs)

---

## Search APIs

### Full-Text Message Search

**Endpoint:** `POST https://substrate.office.com/searchservice/api/v2/query`

**Auth:** Bearer (Substrate)

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
- Response includes `Total` count

**Search Operators:**

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:john.smith@company.com` | Messages from a person |
| `in:` | `in:general` | Messages in a channel |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | By date (explicit dates only) |
| `subject:` | `subject:budget` | In message subject |
| `"Name"` | `"Smith, John"` | Find @mentions (name in quotes) |
| `hasattachment:true` | - | Messages with files |
| `NOT` | `NOT from:user@co.com` | Exclude results |

**Finding @mentions:**
```
"Macdonald, Rob"              # Find mentions of you
"Macdonald, Rob" from:diego   # Mentions from a specific person
```

---

### People Search (Autocomplete)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar`

**Auth:** Bearer (Substrate)

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

### Frequent Contacts

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache`

**Auth:** Bearer (Substrate)

Same as people search, but with empty `QueryString`. Returns ranked list of frequently contacted people.

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

---

### Channel Search (Organisation-wide)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel`

**Auth:** Bearer (Substrate)

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
          "PropertyHits": ["Name"]
        }
      ]
    }
  ]
}
```

**Key Fields:**
- `Name`: Channel display name
- `ThreadId`: Conversation ID for use with messaging APIs
- `TeamName`: Parent team's display name
- `GroupId`: Team's Azure AD group ID
- `ChannelType`: `"Standard"`, `"Private"`, or `"Shared"`

---

## People APIs

### Person Profile (Delve)

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v2/person?smtp={email}&personaType=User&locale=en-gb`

**Auth:** Bearer

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

### Batch User Lookup

**Endpoint:** `POST https://teams.microsoft.com/api/mt/part/{region}/beta/users/fetch`

**Auth:** Bearer

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

### Profile Picture

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/users/{userId}/profilepicturev2/{mri}?size=HR96x96`

Available sizes: `HR64x64`, `HR96x96`, `HR196x196`

---

## Virtual Conversations

Teams uses special "virtual conversation" IDs that act as aggregated views across all conversations. These follow the same messaging API pattern but return consolidated data.

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{virtualId}/messages?view=msnp24Equivalent&pageSize=200&startTime=1`

**Auth:** Skype Token + Bearer

### Available Virtual Conversations

| Virtual ID | Purpose | Notes |
|------------|---------|-------|
| `48:saved` | Saved/bookmarked messages | Messages you've bookmarked across all conversations |
| `48:threads` | Followed threads | Threads you're following for updates |
| `48:mentions` | @mentions | Messages where you were @mentioned |
| `48:notifications` | Activity feed | All notifications (mentions, reactions, replies) |
| `48:notes` | Personal notes | Self-chat / notes to self |
| `48:drafts` | Draft messages | Unsent scheduled messages (different endpoint pattern) |

### Response Structure

Virtual conversation messages include additional fields to identify the source:

```json
{
  "messages": [
    {
      "sequenceId": 55,
      "conversationid": "48:saved",
      "conversationLink": "https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations/48:saved",
      "contenttype": "text",
      "type": "Message",
      "s2spartnername": "skypespaces",
      "clumpId": "19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2",
      "secondaryReferenceId": "T_19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2_M_1769464929223",
      "id": "1769470012345",
      "originalarrivaltime": "2026-01-26T18:30:00.000Z",
      "content": "Message content here...",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John"
    }
  ]
}
```

**Key Fields:**

| Field | Description |
|-------|-------------|
| `clumpId` | The original conversation ID where the message lives |
| `secondaryReferenceId` | Composite key: `T_{conversationId}_M_{messageId}` for messages, `T_{conversationId}_P_{postId}_Threads` for followed threads |
| `id` | Message ID within the virtual conversation (not the original message ID) |
| `originalarrivaltime` | Original timestamp from source conversation |

### Drafts Endpoint (Different Pattern)

Drafts use a slightly different endpoint:

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/drafts?view=msnp24Equivalent&pageSize=200&startTime=1`

**Response:**
```json
{
  "drafts": [
    {
      "sequenceId": 1,
      "conversationid": "48:drafts",
      "draftType": "ScheduledDraft",
      "innerThreadId": "19:abc_def@unq.gbl.spaces",
      "draftDetails": {
        "sendAt": "1755475200000"
      },
      "content": "Scheduled message content..."
    }
  ]
}
```

---

## Messaging APIs

### Send Message

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages`

**Auth:** Skype Token + Bearer

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

**Special Conversation IDs:** See [Virtual Conversations](#virtual-conversations) for the full list (`48:notes`, `48:saved`, `48:threads`, etc.)

---

### Get Messages

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages?pageSize=50&view=msnp24Equivalent`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "messages": [
    {
      "id": "1769189921704",
      "originalarrivaltime": "2026-01-23T17:54:43.263Z",
      "messagetype": "RichText/Html",
      "contenttype": "text",
      "content": "<p>Message content</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John"
    }
  ]
}
```

---

### Reply to Thread (Channel)

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{channelId};messageid={threadRootId}/messages`

**Auth:** Skype Token + Bearer

The `;messageid=` suffix indicates a thread reply. URL-encoded as `%3Bmessageid%3D`.

**URL Pattern Differences:**

| Action | URL Path |
|--------|----------|
| New channel post | `conversations/{channelId}/messages` |
| Reply to thread | `conversations/{channelId};messageid={threadRootId}/messages` |
| Chat message | `conversations/{chatId}/messages` |

**Note:** Chats (1:1, group, meeting) don't use threading. All messages go to the flat conversation.

---

### Edit Message

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}`

**Auth:** Skype Token + Bearer

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

You can only edit your own messages. Returns `403 Forbidden` for others' messages.

---

### Delete Message

**Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}?behavior=softDelete`

**Auth:** Skype Token + Bearer

**Response:** `200 OK` with `null` body

This is a soft delete. Channel owners/moderators can delete others' messages.

---

### Typing Indicator

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

## Conversation APIs

### Create Group Chat

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/threads`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "members": [
    { "id": "8:orgid:user1-guid", "role": "Admin" },
    { "id": "8:orgid:user2-guid", "role": "Admin" },
    { "id": "8:orgid:user3-guid", "role": "Admin" }
  ],
  "properties": {
    "threadType": "chat",
    "topic": "Optional chat name"
  }
}
```

**Response (201):**
```json
{
  "threadResource": {
    "id": "19:5bf2c81dc44b4a60a181bf9170953912@thread.v2",
    "tenantId": "56b731a8-...",
    "type": "Thread",
    "properties": {
      "creator": "8:orgid:user1-guid",
      "threadType": "chat",
      "historydisclosed": "false"
    }
  }
}
```

**Notes:**
- All members get `"role": "Admin"` for group chats
- The `topic` property is optional - sets the chat name
- **Response body may be empty `{}`** - extract conversation ID from `Location` header instead
- Location header format: `https://amer.ng.msg.teams.microsoft.com/v1/threads/19:xxx@thread.v2`
- Use the extracted ID with the messages endpoint to send messages

---

### Get User's Teams & Channels

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true`

**Auth:** Skype Token + Bearer (CSA)

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
            "createdTime": 1753172521981
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

---

### Get Conversation Details

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}?view=msnp24Equivalent`

**Auth:** Skype Token + Bearer

**Response:**
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

---

### Channel Posts

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/containers/{containerId}/posts`

**Query Parameters:**
- `threadedPostsOnly=true` - Only top-level posts
- `pageSize=20`
- `teamId={teamId}`
- `includeRcMetadata=true` - Include read/saved metadata

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
        "content": "<p>Message content</p>",
        "fromFamilyNameInToken": "Smith",
        "fromGivenNameInToken": "John"
      }
    }
  ]
}
```

---

### Favourites (Get/Modify)

**Endpoint:** `POST https://teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true`

**Auth:** Skype Token + Bearer (CSA)

**Get all folders:**
```json
{}
```

**Add to Favorites:**
```json
{
  "actions": [
    {
      "action": "AddItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "{conversationId}"
    }
  ]
}
```

**Remove from Favorites:**
```json
{
  "actions": [
    {
      "action": "RemoveItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "{conversationId}"
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
      "name": "Favorites",
      "folderType": "Favorites",
      "conversationFolderItems": [
        {
          "conversationId": "19:abc@thread.tacv2",
          "createdTime": 1750768187119
        }
      ]
    }
  ]
}
```

---

### Save/Unsave Message (Bookmarks)

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/rcmetadata/{rootMessageId}`

**Auth:** Skype Token + Bearer

**Two-ID System:**

The rcmetadata API uses two different message IDs:
- **URL path** (`rootMessageId`): The thread root post ID for channel threaded replies, or the message ID itself for top-level posts
- **Body** (`mid`): The actual message being saved/unsaved

For **1:1 chats, group chats, meetings, and channel top-level posts**: `rootMessageId` = `messageId` (same value)

For **channel threaded replies**: `rootMessageId` = parent post ID ≠ `messageId`

**Save:**
```json
{ "s": 1, "mid": 1769200192761 }
```

**Unsave:**
```json
{ "s": 0, "mid": 1769200192761 }
```

**Response:**
```json
{
  "conversationId": "19:abc@thread.v2",
  "rootMessageId": 1769200192761,
  "rcMetadata": {
    "lu": 1769200800298,
    "s": 1
  }
}
```

---

### Mark as Read

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/properties?name=consumptionhorizon`

**Auth:** Skype Token + Bearer

**Request:**
```json
{ "consumptionhorizon": "{timestamp1};{timestamp2};{messageId}" }
```

---

### Get Read Position

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/consumptionhorizons`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "id": "19:meeting_abc@thread.v2",
  "version": "1769191217184",
  "consumptionhorizons": []
}
```

---

### 1:1 Chat ID Format

Conversation IDs for 1:1 chats are **predictable** - no API call needed:

```
19:{userId1}_{userId2}@unq.gbl.spaces
```

- User IDs are Azure AD object IDs (GUIDs)
- IDs are **sorted lexicographically** (both participants get the same ID)
- The conversation is created implicitly when the first message is sent

**Example:**
- Your ID: `ab76f827-27e2-4c67-a765-f1a53145fa24`
- Other: `5817f485-f870-46eb-bbc4-de216babac62`
- Since `'5' < 'a'`: `19:5817f485-..._ab76f827-...@unq.gbl.spaces`

---

## Activity & Notifications

### Activity Feed

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Anotifications/messages?view=msnp24Equivalent&pageSize=50`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "messages": [
    {
      "id": "1769276832046",
      "originalarrivaltime": "2026-01-24T18:47:12.046Z",
      "messagetype": "RichText/Html",
      "content": "<p>Activity content here</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John",
      "conversationid": "19:meeting_abc123@thread.v2",
      "threadtopic": "Weekly Standup"
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

Use `syncState` from response for efficient incremental polling.

---

### Thread Annotations (Reactions, Read Status)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/annotations?messageIds={id1},{id2}`

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

## Reactions & Emoji

### Add Reaction

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "emotions": {
    "key": "like",
    "value": 1769429691997
  }
}
```

### Remove Reaction

**Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Request:**
```json
{
  "emotions": {
    "key": "like"
  }
}
```

### Emoji Key Format

**Standard emojis:** Just the name (e.g., `like`, `heart`, `laugh`)

**Custom/org emojis:** `{name};{storage-id}` (e.g., `octi-search;0-wus-d10-66fac2a3b0cda332435c21a14485efe7`)

**Quick Reaction Keys:**

| Emoji | Key |
|-------|-----|
| 👍 | `like` |
| ❤️ | `heart` |
| 😂 | `laugh` |
| 😮 | `surprised` |
| 😢 | `sad` |
| 😠 | `angry` |

**Other Common Keys:**

| Category | Keys |
|----------|------|
| Expressions | `smile`, `wink`, `cry`, `cwl`, `rofl`, `blush`, `speechless`, `wonder`, `sleepy`, `yawn`, `eyeroll`, `worry`, `puke`, `giggle` |
| Affection | `kiss`, `inlove`, `hug`, `lips` |
| Actions | `facepalm`, `sweat`, `dance`, `bow`, `headbang`, `wasntme`, `hungover`, `shivering` |
| Animals | `penguin`, `cat`, `monkey`, `polarbear`, `elephant` |
| Objects | `flower`, `sun`, `star`, `xmastree`, `cake`, `gift`, `cash`, `champagne` |

---

### Custom Emoji Metadata

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/customemoji/metadata`

**Response:**
```json
{
  "categories": [
    {
      "id": "customEmoji",
      "title": "Custom Emoji",
      "emoticons": [
        {
          "id": "angrysteam;0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "documentId": "0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "shortcuts": ["angrysteam"],
          "description": "angrysteam"
        }
      ]
    }
  ]
}
```

**Image URL pattern:** `https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/{emoji-id}/default/20_f.png`

---

## Calendar & Scheduling

### Schedule / Availability

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/schedule?smtp={email}&personaType=User`

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
    "timeZone": { "name": "Greenwich Mean Time" }
  }
}
```

---

### Out of Office Status

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/oofstatus?smtp={email}&personaType=User`

**Response:**
```json
{
  "outOfOfficeState": "Disabled",
  "externalAudience": "Unknown",
  "emailAddress": "USER@COMPANY.COM"
}
```

---

### Meeting Tabs

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/chats/{meetingThreadId}/tabs`

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

## Files & Attachments

### Files Shared in Conversation

**Endpoint:** `GET https://substrate.office.com/AllFiles/api/users('OID:{userId}@{tenantId}')/AllShared?ThreadId={conversationId}&ItemTypes=File&ItemTypes=Link&PageSize=25`

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

## Common Gotchas

1. **Date operators** - Only explicit dates work (`sent:2026-01-20`) or `sent:today`. Named shortcuts like `sent:lastweek` and `sent:thisweek` return 0 results.

2. **`@me` doesn't exist** - `from:me`, `to:me`, and `mentions:me` don't work. Get your email/name first, then search with those values.

3. **Thread replies require `;messageid=`** - The URL suffix is required for channel thread replies. Chats don't have threading.

4. **Token expiry** - MSAL tokens last ~1 hour. They only refresh when an API call requires them, not on page load.

5. **CSA vs chatsvc auth** - CSA needs the CSA Bearer token; chatsvc uses the skypetoken cookie.

6. **Search won't find all thread replies** - It's full-text search. A reply that doesn't contain your search terms won't appear. Use `teams_get_thread` for full context.

7. **User ID formats vary** - APIs return IDs as raw GUIDs, MRIs (`8:orgid:...`), with tenant suffixes (`...@tenantId`), or base64-encoded. Handle all formats.

8. **Message deep links** - Format: `https://teams.microsoft.com/l/message/{threadId}/{messageTimestamp}`. Use the thread ID, not the channel ID, for threaded messages.

---

## Conducting API Research

To discover new endpoints:

1. Clear session: `rm -rf ~/.msteams-mcp/` (or `%APPDATA%\msteams-mcp\` on Windows)
2. Run: `npm run research`
3. Log in to Teams when prompted
4. Perform actions (click Activity, Calendar, etc.)
5. Press Ctrl+C to stop
6. Check terminal output for captured requests/responses

For initial boot capture, clear all Teams data before starting the research script.
