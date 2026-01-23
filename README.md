# Teams Chat Export

Export Microsoft Teams chat messages to Markdown format.

## Features

- Captures sender names and timestamps (using ISO dates for accuracy)
- Preserves links (formatted as markdown links)
- Captures emoji reactions
- Detects edited messages
- **Expands and captures thread replies** (optional, can be disabled)
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
6. A modal will appear with:
   - Detected chat/channel name
   - Days to capture (default: 2)
   - Checkbox to expand threads (enabled by default)
7. Click **Export**
8. Wait for the progress bar - it will:
   - Scan the main chat
   - Expand each thread to capture replies (if enabled)
9. Markdown is copied to your clipboard
10. Paste wherever you need it

## Options

### Days to Capture
How many days back to include. Messages older than this are filtered out.

### Expand Thread Replies
When enabled, the script will:
1. Find all messages with thread replies
2. Click each one to open the thread panel
3. Extract all reply messages
4. Include them in the output nested under the parent message

This is slower but gives you complete thread content. Disable if you only need the main chat messages.

## Output Format

```markdown
# Project Alpha Team

**Exported:** 15/03/2025

---

## Monday 14 March 2025

**Smith, Jane** (09:15):
> Morning all! Quick reminder that the design review is at 2pm today.
>
> 🔗 [Meeting link](https://example.com/meeting)
>
> 3 Like reactions
>
> 💬 **Thread (2 replies):**
>
> > **Chen, David** (09:22):
> > > Thanks for the reminder! I'll be there.
>
> > **Williams, Sarah** (09:45):
> > > Running 5 mins late but on my way

**Chen, David** (10:30) *(edited)*:
> Just pushed the latest changes to the feature branch. Ready for review when you have a moment.
>
> 🔗 [Pull request](https://github.com/example/repo/pull/123)

---

## Tuesday 15 March 2025

**Williams, Sarah** (08:45):
> Has anyone seen the updated requirements doc?
>
> 💬 **Thread (3 replies):**
>
> > **Smith, Jane** (08:52):
> > > I uploaded it to the shared drive yesterday
> > >
> > > 1 Like reaction
>
> > **Chen, David** (09:01):
> > > Found it, thanks!
>
> > **Williams, Sarah** (09:05):
> > > Perfect, got it now
```

## How It Works

1. **Scrolling**: Teams uses virtual scrolling (only visible messages are in the DOM). The script scrolls through the chat to load and capture all messages.

2. **ISO Dates**: Uses the `datetime` attribute on timestamp elements for accurate sorting, not the displayed text.

3. **Thread Detection**: Looks for `replies-summary-authors` elements to identify messages with threads.

4. **Thread Expansion**: Clicks thread buttons, waits for the right-rail panel to load, extracts replies, then closes the panel.

5. **Preview Filtering**: Messages starting with "Replied in thread:" are preview messages, not originals. These are filtered out to avoid duplication.

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
Check the browser console (F12 → Console) where the markdown is also logged. Copy it from there.

### TrustedHTML error
This occurs if you try to use `innerHTML`. The provided script uses `createElement` to avoid this.

## Technical Notes

- Works with Teams web app (teams.microsoft.com)
- Tested with the new Teams interface (Fluent UI)
- DOM selectors used:
  - `chat-pane-list` - Main message container
  - `chat-pane-item` - Individual message wrapper
  - `chat-pane-message` - Message content area
  - `message-author-name` - Sender name
  - `replies-summary-authors` - Thread indicator
  - `right-rail-message-pane-body` - Thread panel

## Files

- `teams-export.js` - The full export script (run in browser console)
- `README.md` - This documentation
- `examples/` - Your own exports (gitignored, not committed)
