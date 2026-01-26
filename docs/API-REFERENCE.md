# Microsoft Teams API Reference

Quick reference for the internal Teams APIs used by this MCP server. For detailed research notes, see [API-RESEARCH.md](./API-RESEARCH.md).

## Authentication

| Auth Type | Header | Source |
|-----------|--------|--------|
| **Bearer (Substrate)** | `Authorization: Bearer {token}` | MSAL localStorage, `SubstrateSearch` scope |
| **Bearer (CSA)** | `Authorization: Bearer {csaToken}` | MSAL, `chatsvcagg.teams.microsoft.com` audience |
| **Skype Token** | `Authentication: skypetoken={token}` | Cookie `skypetoken_asm` |

Most endpoints also require:
```
Origin: https://teams.microsoft.com
Referer: https://teams.microsoft.com/
Content-Type: application/json
```

---

## Search APIs

### Search Messages

```
POST https://substrate.office.com/searchservice/api/v2/query
Auth: Bearer (Substrate)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `entityRequests[].query.queryString` | string | Search query with operators |
| `entityRequests[].from` | number | Offset for pagination (0, 25, 50...) |
| `entityRequests[].size` | number | Page size (max ~50) |

**Search Operators:**
- `from:email@company.com` — Messages from person
- `in:channel-name` — Messages in channel
- `sent:2026-01-20` or `sent:>=2026-01-15` — By date
- `"Display Name"` — Find @mentions
- `hasattachment:true` — Has files
- `NOT term` — Exclude results

**Response:** `EntitySets[0].ResultSets[0].Results[]` with `Total` count for pagination.

---

### Search People

```
POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar
Auth: Bearer (Substrate)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `EntityRequests[].Query.QueryString` | string | Name or email to search |
| `EntityRequests[].EntityType` | string | `"People"` |
| `EntityRequests[].Size` | number | Max results |

**Response:** `Groups[0].Suggestions[]` with `DisplayName`, `EmailAddresses`, `JobTitle`, etc.

---

### Frequent Contacts

```
POST https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache
Auth: Bearer (Substrate)
```

Same as people search, but with empty `QueryString`. Returns ranked list of frequently contacted people.

---

### Search Channels (Org-wide)

```
POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel
Auth: Bearer (Substrate)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `EntityRequests[].Query.QueryString` | string | Channel name to search |
| `EntityRequests[].EntityType` | string | `"TeamsChannel"` |
| `cvid` | string | Correlation ID (UUID) |
| `logicalId` | string | Logical ID (UUID) |

**Response:** `Groups[0].Suggestions[]` with `Name`, `ThreadId`, `TeamName`, `GroupId`, `ChannelType`.

---

## Messaging APIs

### Get Messages

```
GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages
Auth: Skype Token + Bearer
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `pageSize` | number | Messages to return (query param) |
| `view` | string | `msnp24Equivalent` (query param) |

**Response:** `messages[]` with `id`, `content`, `from`, `originalarrivaltime`, `imdisplayname`.

---

### Send Message

```
POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages
Auth: Skype Token + Bearer
```

| Field | Type | Description |
|-------|------|-------------|
| `content` | string | Message HTML content |
| `messagetype` | string | `"RichText/Html"` |
| `clientmessageid` | string | Unique client ID |

**Special IDs:** `48:notes` (self-chat), `48:notifications` (activity feed).

---

### Reply to Thread

```
POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{channelId};messageid={threadRootId}/messages
Auth: Skype Token + Bearer
```

Same body as send message. The `;messageid=` suffix indicates thread reply.

---

### Edit Message

```
PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}
Auth: Skype Token + Bearer
```

| Field | Type | Description |
|-------|------|-------------|
| `content` | string | Updated message content |
| `messagetype` | string | `"RichText/Html"` |

Only works for your own messages.

---

### Delete Message

```
DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}?behavior=softDelete
Auth: Skype Token + Bearer
```

Soft delete only. Returns `200 OK` with `null` body.

---

## Conversation APIs

### Get User's Teams & Channels

```
GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true
Auth: Skype Token + Bearer (CSA)
```

**Response:** `teams[]` with nested `channels[]`. Each channel has `id`, `displayName`, `membershipType`.

---

### Get Conversation Details

```
GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}?view=msnp24Equivalent
Auth: Skype Token + Bearer
```

**Response:** `threadProperties` with `threadType`, `productThreadType`, `topic`, etc.

| threadType | productThreadType | Type |
|------------|-------------------|------|
| `topic` | `TeamsStandardChannel` | Channel |
| `space` | `TeamsTeam` | Team root |
| `space` | `TeamsPrivateChannel` | Private channel |
| `meeting` | `Meeting` | Meeting chat |
| `chat` | `Chat` | Group chat |
| `chat` | `OneOnOne` | 1:1 chat |

---

### Favourites (Get/Modify)

```
POST https://teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true
Auth: Skype Token + Bearer (CSA)
```

**Get all:** Send empty `{}` body.

**Add/Remove:**
```json
{
  "actions": [
    { "action": "AddItem", "folderId": "{tenantId}~{userId}~Favorites", "itemId": "{conversationId}" }
  ]
}
```

---

### Save/Unsave Message

```
PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/rcmetadata/{messageId}
Auth: Skype Token + Bearer
```

```json
{ "s": 1, "mid": 1769200192761 }
```

`s: 1` = save, `s: 0` = unsave.

---

### Mark as Read

```
PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/properties?name=consumptionhorizon
Auth: Skype Token + Bearer
```

```json
{ "consumptionhorizon": "{timestamp1};{timestamp2};{messageId}" }
```

---

### Get Read Position

```
GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/consumptionhorizons
Auth: Skype Token + Bearer
```

---

## Activity Feed

### Get Activity

```
GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Anotifications/messages?view=msnp24Equivalent&pageSize=50
Auth: Skype Token + Bearer
```

**Response:** `messages[]` — mentions, reactions, replies. Check `messagetype` and content patterns to identify activity type.

---

## People APIs

### Person Profile

```
POST https://nam.loki.delve.office.com/api/v2/person?smtp={email}&personaType=User
Auth: Bearer
```

**Response:** `person` with `names`, `emailAddresses`, `jobTitle`, etc.

---

### Batch User Lookup

```
POST https://teams.microsoft.com/api/mt/part/{region}/beta/users/fetch
Auth: Bearer
```

```json
["8:orgid:{userId1}", "8:orgid:{userId2}"]
```

---

## Regional Codes

| Region | Code |
|--------|------|
| Americas | `amer` |
| Europe/Middle East/Africa | `emea` |
| Asia Pacific | `apac` |

Used in: `/api/csa/{region}/`, `/api/chatsvc/{region}/`, `/api/mt/part/{region}/`

---

## 1:1 Chat ID Format

Conversation IDs for 1:1 chats are predictable:

```
19:{userId1}_{userId2}@unq.gbl.spaces
```

The two user object IDs (GUIDs) are sorted lexicographically. No API call needed — the conversation is created when the first message is sent.

---

## Common Gotchas

1. **Date operators** — Only explicit dates work (`sent:2026-01-20`). Named shortcuts like `sent:lastweek` return 0 results.

2. **`@me` doesn't exist** — Use `teams_get_me` to get your email/name, then search with those values.

3. **Thread replies** — The `;messageid=` URL suffix is required for channel thread replies. Chats don't have threading.

4. **Token expiry** — MSAL tokens last ~1 hour. Proactive refresh triggers via API call, not page load.

5. **CSA vs chatsvc** — Different APIs need different auth. CSA needs the CSA token; chatsvc uses skypetoken.
