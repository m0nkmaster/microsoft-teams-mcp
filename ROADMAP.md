# Future plans and bugs to fix

## ✅ find channel

Implemented `teams_find_channel` tool. Uses the Substrate suggestions API with `domain=TeamsChannel` to search ALL channels across the organisation (not just channels in the user's teams). Returns channel name, team name, and conversation ID for use with `teams_get_thread`.

## Token refresh mechanism

Proactively refresh tokens before they expire (~1 hour) instead of falling back to browser login.

## find team

Search/discover teams by name (similar to find channel but returns team-level info).

## meeting related stuff

- Get messages from meeting chat threads
- Calendar integration (upcoming meetings)
- Find meeting recordings
