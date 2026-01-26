# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| P1 | Login/refresh UX | Show visual feedback when auth succeeds before browser auto-closes | Easy | Playwright injection |
| P2 | Find team | Search/discover teams by name | Easy | Teams List API |
| P2 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API |
| P3 | Calendar integration | Get upcoming meetings | Hard | Outlook API, needs research |
| P3 | Meeting recordings | Locate recordings/transcripts | Hard | Needs research |
| P3 | Region auto-detection | Detect user's API region (amer/emea/apac) from session instead of defaulting to amer | Easy | Could extract from browser session or make configurable via env var |