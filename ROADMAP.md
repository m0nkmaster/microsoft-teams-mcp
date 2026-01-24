# Future plans and bugs to fix

## Token refresh mechanism

Proactively refresh tokens before they expire (~1 hour) instead of falling back to browser login.

## find team

Search/discover teams by name (similar to find channel but returns team-level info).

## Improve UI for login/token refresh (popup message to close browser)

## meeting related stuff

- Get messages from meeting chat threads
- Calendar integration (upcoming meetings)
- Find meeting recordings


## ✅ Edit and Delete messages

Implemented `teams_edit_message` and `teams_delete_message` tools. Edit uses PUT to update message content; delete uses DELETE with `?behavior=softDelete` query parameter. Both only work on your own messages (unless you're a channel moderator for delete).
