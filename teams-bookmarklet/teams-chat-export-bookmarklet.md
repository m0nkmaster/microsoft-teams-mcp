# Teams Chat Export

Export Microsoft Teams chat messages to Markdown format.

## Features

- Captures sender names and timestamps (using ISO dates for accuracy)
- Preserves links (formatted as markdown links)
- Captures emoji reactions
- Detects edited messages
- **Expands and captures thread replies** (optional, can be disabled)
- **Detects open threads** and offers to export just the thread
- Filters out "Replied in thread" preview messages
- Sorts messages chronologically
- Filters by configurable date range (days back)
- Groups messages by date with section headers

## Usage

### Console Script (Recommended)

Due to Teams' strict Content Security Policy, bookmarklets are blocked. Use the console script instead:

1. Open Teams in your browser (teams.microsoft.com)
2. Navigate to the chat/channel you want to export
3. Open DevTools (F12) → Console tab
4. Copy the contents of `teams-export.js` and paste into the console
5. Press Enter

### If a Thread is Open

If you have a thread panel open when you run the script, you'll see:

```
┌─────────────────────────────────────┐
│ Thread Detected                     │
│                                     │
│ ⚠️ A thread is currently open with  │
│ X messages.                         │
│                                     │
│ What would you like to export?      │
│                                     │
│ [💬 Export This Thread Only]        │
│ [📋 Export Full Chat]               │
│ [Cancel]                            │
└─────────────────────────────────────┘
```

- **Export This Thread Only** - Immediately exports just the thread messages
- **Export Full Chat** - Closes the thread and shows the full chat export options

### Normal Export (No Thread Open)

You'll see the standard export dialog:

1. **Chat name** - Detected automatically from the page
2. **Days to capture** - How many days back to include (default: 2)
3. **Expand thread replies** - When enabled, clicks each thread to capture replies
4. Click **Export**
5. Wait for the progress bar
6. Markdown is copied to your clipboard

## How It Works

1. **Thread Detection**: Checks if `right-rail-message-pane-body` is visible when script starts

2. **Scrolling**: Teams uses virtual scrolling (only visible messages are in the DOM). The script scrolls through the chat to load and capture all messages.

3. **ISO Dates**: Uses the `datetime` attribute on timestamp elements for accurate sorting, not the displayed text.

4. **Thread Expansion**: Clicks thread buttons, waits for the right-rail panel to load, extracts replies, then closes the panel.

5. **Preview Filtering**: Messages starting with "Replied in thread:" are preview messages, not originals. These are filtered out.

6. **DOM Creation**: Uses `document.createElement` instead of `innerHTML` to bypass Teams' Content Security Policy.

## Troubleshooting

### "Chat pane not found"
Make sure you're viewing an active chat/channel, not the chat list or settings.

### Not all messages captured
For very long chats, the scroll might not capture everything. Try:
- Running it twice
- Using a smaller "days to capture" value

### Threads not expanding
Some threads might not expand if:
- The message scrolled out of view before clicking
- Teams' virtual DOM recycled the element

Check the browser console for messages like `Could not find thread:` to see which ones were missed.

### Clipboard access denied
Check the browser console (F12 → Console) where the markdown is also logged.

## Supported Scenarios

| Scenario | Status | Notes |
|----------|--------|-------|
| Channel chat | ✅ Works | Full support |
| DM chat | ✅ Works | Same DOM structure |
| Meeting chat | ✅ Works | Same DOM structure |
| Thread open | ✅ Works | Offers to export just thread |
| Group chat | ✅ Works | Same DOM structure |

## Files

- `teams-export.js` - The full export script (run in browser console)
- `teams-chat-export-bookmarklet.md` - This documentation
