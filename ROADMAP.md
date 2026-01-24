# Future plans and bugs to fix

## ✅ find channel

Implemented `teams_find_channel` tool. Uses the CSA v3 `/api/csa/{region}/api/v3/teams/users/me` endpoint to fetch all teams and channels, then filters by name.

## Token refresh mechanism

Proactively refresh tokens before they expire (~1 hour) instead of falling back to browser login.

## find team

Search/discover teams by name (similar to find channel but returns team-level info).

## meeting related stuff

- Get messages from meeting chat threads
- Calendar integration (upcoming meetings)
- Find meeting recordings
