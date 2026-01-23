/**
 * Teams Chat Export Script
 * 
 * Exports Microsoft Teams chat messages to Markdown format.
 * Run this in the browser console while viewing a Teams chat.
 * 
 * Features:
 * - Captures sender, timestamp, content
 * - Preserves links and reactions
 * - Detects edited messages
 * - Expands and captures thread replies
 * - Filters out "Replied in thread" preview messages
 * - Sorts chronologically using ISO dates
 * - Filters by configurable date range
 * 
 * Usage:
 * 1. Open Teams in browser (teams.microsoft.com)
 * 2. Navigate to the chat/channel to export
 * 3. Open DevTools (F12) → Console tab
 * 4. Paste this entire script and press Enter
 * 5. Configure days to capture and thread expansion
 * 6. Click Export
 */
(async () => {
  const chatTitle = document.querySelector('h2')?.textContent || 'Teams Chat';
  
  // Create overlay
  const overlay = document.createElement('div');
  Object.assign(overlay.style, {
    position: 'fixed', top: '0', left: '0', right: '0', bottom: '0',
    background: 'rgba(0,0,0,0.5)', zIndex: '999998'
  });
  
  // Create modal
  const modal = document.createElement('div');
  Object.assign(modal.style, {
    position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%,-50%)',
    background: '#fff', padding: '24px', borderRadius: '12px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.3)', zIndex: '999999',
    fontFamily: 'system-ui', minWidth: '400px'
  });
  
  const title = document.createElement('h2');
  title.textContent = 'Export Teams Chat';
  Object.assign(title.style, { margin: '0 0 16px' });
  
  const info = document.createElement('div');
  info.textContent = 'Chat: ' + chatTitle;
  Object.assign(info.style, {
    background: '#f5f5f5', padding: '12px', borderRadius: '6px', marginBottom: '16px'
  });
  
  const label = document.createElement('label');
  label.textContent = 'Days to capture: ';
  Object.assign(label.style, { display: 'block', marginBottom: '8px' });
  
  const input = document.createElement('input');
  input.type = 'number';
  input.value = '2';
  input.min = '1';
  input.max = '30';
  Object.assign(input.style, {
    width: '100%', padding: '8px', border: '1px solid #ddd',
    borderRadius: '6px', marginTop: '4px', boxSizing: 'border-box'
  });
  label.appendChild(input);
  
  // Thread expansion checkbox
  const threadLabel = document.createElement('label');
  Object.assign(threadLabel.style, {
    display: 'flex', alignItems: 'center', gap: '8px', marginTop: '12px', cursor: 'pointer'
  });
  const threadCheck = document.createElement('input');
  threadCheck.type = 'checkbox';
  threadCheck.checked = true;
  threadLabel.appendChild(threadCheck);
  threadLabel.appendChild(document.createTextNode('Expand and capture thread replies (slower)'));
  
  const progressArea = document.createElement('div');
  progressArea.style.display = 'none';
  progressArea.style.marginTop = '16px';
  
  const progressBarOuter = document.createElement('div');
  Object.assign(progressBarOuter.style, {
    height: '8px', background: '#e0e0e0', borderRadius: '4px', overflow: 'hidden'
  });
  
  const progressBar = document.createElement('div');
  Object.assign(progressBar.style, {
    height: '100%', background: '#6264a7', width: '0%', transition: 'width 0.3s'
  });
  progressBarOuter.appendChild(progressBar);
  
  const progressText = document.createElement('div');
  Object.assign(progressText.style, { fontSize: '12px', color: '#666', marginTop: '8px' });
  progressText.textContent = 'Starting...';
  
  progressArea.appendChild(progressBarOuter);
  progressArea.appendChild(progressText);
  
  const buttonArea = document.createElement('div');
  Object.assign(buttonArea.style, {
    display: 'flex', gap: '12px', justifyContent: 'flex-end', marginTop: '20px'
  });
  
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel';
  Object.assign(cancelBtn.style, {
    padding: '10px 20px', border: 'none', borderRadius: '6px',
    cursor: 'pointer', background: '#f0f0f0'
  });
  
  const exportBtn = document.createElement('button');
  exportBtn.textContent = 'Export';
  Object.assign(exportBtn.style, {
    padding: '10px 20px', border: 'none', borderRadius: '6px',
    cursor: 'pointer', background: '#6264a7', color: '#fff'
  });
  
  buttonArea.appendChild(cancelBtn);
  buttonArea.appendChild(exportBtn);
  
  modal.appendChild(title);
  modal.appendChild(info);
  modal.appendChild(label);
  modal.appendChild(threadLabel);
  modal.appendChild(progressArea);
  modal.appendChild(buttonArea);
  
  document.body.appendChild(overlay);
  document.body.appendChild(modal);
  
  const close = () => { overlay.remove(); modal.remove(); };
  overlay.onclick = close;
  cancelBtn.onclick = close;

  // Helper to extract messages from a pane (used for thread replies)
  const extractFromPane = (pane) => {
    const msgs = [];
    pane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
      const msg = item.querySelector('[data-tid="chat-pane-message"]');
      const ctrl = item.querySelector('[data-tid="control-message-renderer"]');
      let sender = '', isoDate = '', content = '', edited = false, links = [], reactions = [];
      
      if (ctrl) {
        sender = '[System]';
        content = ctrl.textContent?.trim() || '';
        isoDate = item.querySelector('time')?.getAttribute('datetime') || '';
      } else if (msg) {
        sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
        const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
        isoDate = timeEl?.getAttribute('datetime') || '';
        content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
        edited = !!item.querySelector('[id^="edited-"]');
        links = [...msg.querySelectorAll('a[href]')]
          .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
          .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
        reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
          .map(r => r.textContent?.trim()).filter(Boolean);
      }
      if (content) {
        msgs.push({
          sender, isoDate, content, edited,
          links: links.length ? links : null,
          reactions: reactions.length ? reactions : null
        });
      }
    });
    return msgs;
  };

  // Find a message's thread button by matching content (handles stale DOM references)
  const findThreadButton = (chatPane, contentSnippet, isoDate) => {
    const items = chatPane.querySelectorAll('[data-tid="chat-pane-item"]');
    for (const item of items) {
      const msgContent = item.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
      const msgTime = item.querySelector('time')?.getAttribute('datetime') || '';
      
      if (msgContent.startsWith(contentSnippet.substring(0, 30))) {
        if (!isoDate || msgTime === isoDate) {
          const replySummary = item.querySelector('[data-tid="replies-summary-authors"]');
          if (replySummary) {
            return replySummary.closest('button');
          }
        }
      }
    }
    return null;
  };

  exportBtn.onclick = async () => {
    const days = parseInt(input.value) || 2;
    const expandThreads = threadCheck.checked;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    cutoff.setHours(0, 0, 0, 0);
    
    buttonArea.style.display = 'none';
    progressArea.style.display = 'block';

    const messages = new Map();
    const chatPane = document.getElementById('chat-pane-list');
    if (!chatPane) {
      progressText.textContent = 'Error: No chat pane found. Make sure a chat is open.';
      return;
    }
    
    const viewport = chatPane.parentElement;

    // Extract messages from main chat pane
    const extract = () => {
      chatPane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
        const msg = item.querySelector('[data-tid="chat-pane-message"]');
        const ctrl = item.querySelector('[data-tid="control-message-renderer"]');
        let sender = '', timeDisplay = '', isoDate = '', content = '';
        let edited = false, links = [], reactions = [], threadInfo = null;
        let isThreadPreview = false;
        
        if (ctrl) {
          sender = '[System]';
          content = ctrl.textContent?.trim() || '';
          isoDate = item.querySelector('time')?.getAttribute('datetime') || '';
        } else if (msg) {
          sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
          const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
          timeDisplay = timeEl?.textContent?.trim() || '';
          isoDate = timeEl?.getAttribute('datetime') || '';
          content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
          edited = !!item.querySelector('[id^="edited-"]');
          
          links = [...msg.querySelectorAll('a[href]')]
            .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
            .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
          
          reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
            .map(r => r.textContent?.trim()).filter(Boolean);
          
          // Check if this is a "Replied in thread" preview (not an original message)
          if (content.startsWith('Replied in thread:')) {
            isThreadPreview = true;
            content = content.replace(/^Replied in thread:\s*/, '');
          }
          
          // Check for thread replies - ONLY if replies-summary-authors exists
          const replySummary = item.querySelector('[data-tid="replies-summary-authors"]');
          if (replySummary && !isThreadPreview) {
            const summaryParent = replySummary.closest('[class*="repl"]') || replySummary.parentElement?.parentElement;
            const summaryText = summaryParent?.textContent || '';
            
            // Use word boundary to avoid matching reaction counts
            const replyMatch = summaryText.match(/\b(\d+)\s*repl(?:y|ies)/i);
            if (replyMatch) {
              const lastReplyMatch = summaryText.match(/Last reply\s+([^F]+?)(?:Follow|$)/i);
              threadInfo = {
                replyCount: parseInt(replyMatch[1]),
                lastReply: lastReplyMatch ? lastReplyMatch[1].trim() : null,
                replies: []
              };
            }
          }
        }
        
        if (content) {
          const key = `${sender}-${isoDate || timeDisplay}-${content.substring(0, 40)}`;
          if (!messages.has(key)) {
            messages.set(key, {
              sender, timeDisplay, isoDate, content, edited,
              links: links.length ? links : null,
              reactions: reactions.length ? reactions : null,
              threadInfo,
              isThreadPreview,
              contentSnippet: content.substring(0, 50)
            });
          }
        }
      });
    };

    // Scroll through chat and collect messages
    viewport.scrollTop = viewport.scrollHeight;
    await new Promise(r => setTimeout(r, 500));
    
    let pos = viewport.scrollHeight;
    let done = false;
    
    while (pos > 0 && !done) {
      viewport.scrollTop = pos;
      await new Promise(r => setTimeout(r, 150));
      extract();
      
      progressBar.style.width = Math.round((1 - pos / viewport.scrollHeight) * 40) + '%';
      progressText.textContent = 'Scanning main chat... ' + messages.size + ' messages';
      
      for (const m of messages.values()) {
        if (m.isoDate && new Date(m.isoDate) < cutoff) {
          done = true;
          break;
        }
      }
      pos -= 300;
    }
    extract();

    // Expand threads if enabled
    if (expandThreads) {
      const withThreads = [...messages.values()].filter(m => m.threadInfo && !m.isThreadPreview);
      progressText.textContent = `Found ${withThreads.length} threads to expand...`;
      
      for (let i = 0; i < withThreads.length; i++) {
        const m = withThreads[i];
        progressBar.style.width = (40 + (i / withThreads.length) * 50) + '%';
        progressText.textContent = `Expanding thread ${i + 1}/${withThreads.length}: ${m.sender}...`;
        
        try {
          let threadButton = null;
          let scrollAttempts = 0;
          
          // Try to find thread button in current view
          threadButton = findThreadButton(chatPane, m.contentSnippet, m.isoDate);
          
          // If not found, scroll to find it
          if (!threadButton) {
            viewport.scrollTop = viewport.scrollHeight;
            await new Promise(r => setTimeout(r, 200));
            
            let searchPos = viewport.scrollHeight;
            while (!threadButton && searchPos > 0 && scrollAttempts < 20) {
              viewport.scrollTop = searchPos;
              await new Promise(r => setTimeout(r, 150));
              threadButton = findThreadButton(chatPane, m.contentSnippet, m.isoDate);
              searchPos -= 500;
              scrollAttempts++;
            }
          }
          
          if (threadButton) {
            threadButton.scrollIntoView({ block: 'center' });
            await new Promise(r => setTimeout(r, 200));
            threadButton.click();
            await new Promise(r => setTimeout(r, 1000));
            
            // Extract from right rail thread panel
            const rightRail = document.querySelector('[data-tid="right-rail-message-pane-body"]');
            if (rightRail) {
              await new Promise(r => setTimeout(r, 500));
              const replies = extractFromPane(rightRail);
              // Skip first message (parent) and add rest as replies
              if (replies.length > 1) {
                m.threadInfo.replies = replies.slice(1);
              } else if (replies.length === 1) {
                m.threadInfo.replies = replies;
              }
              console.log(`Thread "${m.contentSnippet.substring(0, 30)}...": ${m.threadInfo.replies.length} replies`);
            }
            
            // Close thread panel
            const toggleBtn = document.querySelector('[data-tid="thread-list-pane-toggle-button"]');
            if (toggleBtn) {
              toggleBtn.click();
              await new Promise(r => setTimeout(r, 400));
            }
          } else {
            console.log(`Could not find thread: ${m.contentSnippet.substring(0, 40)}...`);
          }
        } catch (e) {
          console.log('Error expanding thread:', e);
        }
      }
    }

    // Filter out thread previews and filter by date
    const filtered = [...messages.values()]
      .filter(m => !m.isThreadPreview)
      .filter(m => {
        if (m.isoDate) return new Date(m.isoDate) >= cutoff;
        return true;
      });
    
    // Sort chronologically
    filtered.sort((a, b) => {
      const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
      const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
      return da - db;
    });

    progressBar.style.width = '95%';
    progressText.textContent = 'Generating markdown...';

    // Generate markdown output
    let md = '# ' + chatTitle + '\n\n';
    md += '**Exported:** ' + new Date().toLocaleDateString('en-GB') + '\n\n---\n\n';
    
    let lastDate = '';
    
    for (const m of filtered) {
      let dateStr = '', timeStr = '';
      
      if (m.isoDate) {
        const d = new Date(m.isoDate);
        dateStr = d.toLocaleDateString('en-GB', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
        });
        timeStr = d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      } else {
        dateStr = 'Unknown Date';
        timeStr = m.timeDisplay || '';
      }
      
      // Date header
      if (dateStr !== lastDate) {
        if (lastDate) md += '\n---\n\n';
        md += '## ' + dateStr + '\n\n';
        lastDate = dateStr;
      }
      
      // Message header
      md += '**' + m.sender + '**';
      if (timeStr) md += ' (' + timeStr + ')';
      if (m.edited) md += ' *(edited)*';
      md += ':\n';
      
      // Content
      md += m.content.split('\n').map(l => '> ' + l).join('\n') + '\n';
      
      // Links
      if (m.links) {
        md += '>\n> 🔗 ' + m.links.map(l => '[' + l.text + '](' + l.url + ')').join(' | ') + '\n';
      }
      
      // Reactions
      if (m.reactions) {
        md += '>\n> ' + m.reactions.join(' | ') + '\n';
      }
      
      // Thread replies
      if (m.threadInfo) {
        if (m.threadInfo.replies && m.threadInfo.replies.length > 0) {
          md += '>\n> 💬 **Thread (' + m.threadInfo.replies.length + ' replies):**\n';
          for (const reply of m.threadInfo.replies) {
            let replyTime = '';
            if (reply.isoDate) {
              replyTime = new Date(reply.isoDate).toLocaleTimeString('en-GB', {
                hour: '2-digit', minute: '2-digit'
              });
            }
            md += '>\n> > **' + reply.sender + '**';
            if (replyTime) md += ' (' + replyTime + ')';
            md += ':\n';
            md += reply.content.split('\n').map(l => '> > > ' + l).join('\n') + '\n';
            if (reply.reactions) {
              md += '> > >\n> > > ' + reply.reactions.join(' | ') + '\n';
            }
          }
        } else {
          md += '>\n> 💬 **' + m.threadInfo.replyCount + ' ';
          md += (m.threadInfo.replyCount === 1 ? 'reply' : 'replies') + '**';
          if (m.threadInfo.lastReply) md += ' *(last: ' + m.threadInfo.lastReply + ')*';
          md += ' *[not expanded]*\n';
        }
      }
      
      md += '\n';
    }

    // Copy to clipboard
    try {
      await navigator.clipboard.writeText(md);
      progressBar.style.width = '100%';
      progressText.textContent = '✓ Copied ' + filtered.length + ' messages to clipboard!';
      progressText.style.color = 'green';
      progressText.style.fontWeight = 'bold';
      setTimeout(close, 2500);
    } catch (e) {
      console.log('=== MARKDOWN OUTPUT ===\n', md);
      progressText.textContent = 'Clipboard blocked - check console (F12)';
    }
  };
})();
