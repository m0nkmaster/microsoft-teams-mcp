# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| P1 | Login/refresh UX | Show visual feedback when auth succeeds before browser auto-closes | Easy | Pure Playwright, no API |
| P2 | Find team | Search/discover teams by name | Easy | Teams List API already used by `find_channel` |
| P2 | Add reactions | React to messages with emoji | Medium | `annotations` endpoint discovered, need POST format |
| P2 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API discovered |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API discovered |
| P3 | Calendar integration | Get upcoming meetings | Hard | Outlook APIs exist, need research |
| P3 | Meeting recordings | Locate recordings/transcripts | Hard | Likely Graph API or Stream, need research |