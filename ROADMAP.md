# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| P2 | Create group chat | Create group conversations with multiple people | Medium | ✅ Implemented - `teams_create_group_chat` |
| P2 | Find team | Search/discover teams by name | Easy | Teams List API |
| P2 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API |
| P3 | Calendar integration | Get upcoming meetings | Hard | Outlook API, needs research |
| P3 | Meeting search | Search past meetings by person, topic, date range | Medium | Could use Outlook calendar API or Substrate search with `is:Meetings` filter - needs research |
| P3 | Meeting recordings | Locate recordings/transcripts | Hard | Needs research |
| P3 | Region auto-detection | Detect user's API region (amer/emea/apac) from session instead of defaulting to amer | Easy | Could extract from browser session or make configurable via env var |
| P3 | Verify Skype token refresh | Check whether the messaging token (skypetoken_asm) gets refreshed during auto token refresh, or only the Substrate token | Easy | Add before/after logging to `refreshTokensViaBrowser()` to compare both tokens |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |
| Bug | Message links unreliable | Deep links sometimes fail with "can't find" error when Teams opens | Medium | Investigate link format variations - may be threading/context related |