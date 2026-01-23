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
 * - Detects open threads and offers to export just the thread
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
  
  // Check if a thread is already open
  const rightRail = document.querySelector('[data-tid="right-rail-message-pane-body"]');
  const threadIsOpen = rightRail && rightRail.offsetParent !== null;
  
  // Helper to extract messages from a pane
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

  // Generate markdown from messages array
  const generateMarkdown = (messages, title, isThread = false) => {
    // Sort chronologically
    messages.sort((a, b) => {
      const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
      const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
      return da - db;
    });

    let md = '# ' + title + (isThread ? ' (Thread)' : '') + '\n\n';
    md += '**Exported:** ' + new Date().toLocaleDateString('en-GB') + '\n\n---\n\n';
    
    let lastDate = '';
    
    for (const m of messages) {
      let dateStr = '', timeStr = '';
      
      if (m.isoDate) {
        const d = new Date(m.isoDate);
        dateStr = d.toLocaleDateString('en-GB', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
        });
        timeStr = d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      } else {
        dateStr = 'Unknown Date';
        timeStr = '';
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
      
      // Thread replies (for full chat export)
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
    
    return md;
  };

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
  
  const close = () => { overlay.remove(); modal.remove(); };
  
  // Helper to scroll through a thread pane and extract all messages
  const extractAllFromThread = async (pane) => {
    const msgs = new Map();
    const extract = () => {
      pane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
        const msg = item.querySelector('[data-tid="chat-pane-message"]');
        if (msg) {
          const sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
          const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
          const isoDate = timeEl?.getAttribute('datetime') || '';
          const content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
          const edited = !!item.querySelector('[id^="edited-"]');
          const links = [...msg.querySelectorAll('a[href]')]
            .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
            .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
          const reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
            .map(r => r.textContent?.trim()).filter(Boolean);
          
          if (content) {
            const key = `${sender}-${isoDate}-${content.substring(0, 40)}`;
            if (!msgs.has(key)) {
              msgs.set(key, {
                sender, isoDate, content, edited,
                links: links.length ? links : null,
                reactions: reactions.length ? reactions : null
              });
            }
          }
        }
      });
    };
    
    // Find the scrollable container - walk up the DOM looking for scrollable element
    let scrollContainer = pane;
    let el = pane;
    while (el && el !== document.body) {
      if (el.scrollHeight > el.clientHeight + 10) {
        scrollContainer = el;
        break;
      }
      el = el.parentElement;
    }
    console.log('Thread scroll container:', scrollContainer.className, 'scrollHeight:', scrollContainer.scrollHeight, 'clientHeight:', scrollContainer.clientHeight);
    
    // Scroll to BOTTOM first (threads show newest at bottom, we need to scroll UP to get older)
    scrollContainer.scrollTop = scrollContainer.scrollHeight;
    await new Promise(r => setTimeout(r, 300));
    extract();
    console.log('After initial extract:', msgs.size, 'messages');
    
    // Scroll UP through thread to load older messages
    let scrollAttempts = 0;
    const maxScrollAttempts = 50;
    while (scrollAttempts < maxScrollAttempts) {
      const prevCount = msgs.size;
      const prevScroll = scrollContainer.scrollTop;
      
      scrollContainer.scrollTop -= 400;
      await new Promise(r => setTimeout(r, 200));
      extract();
      
      // Stop if we've hit the top and no new messages
      if (msgs.size === prevCount && (scrollContainer.scrollTop === prevScroll || scrollContainer.scrollTop === 0)) {
        break;
      }
      scrollAttempts++;
    }
    
    console.log(`Thread extraction: ${msgs.size} messages (scrolled ${scrollAttempts}x)`);
    
    // Sort by timestamp
    return [...msgs.values()].sort((a, b) => {
      const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
      const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
      return da - db;
    });
  };
  
  // If thread is open, show thread export option
  if (threadIsOpen) {
    // Get initial count (visible messages only, for display)
    const visibleCount = rightRail.querySelectorAll('[data-tid="chat-pane-item"]').length;

    const title = document.createElement('h2');
    title.textContent = 'Thread Detected';
    Object.assign(title.style, { margin: '0 0 16px', color: '#242424' });
    
    const info = document.createElement('div');
    Object.assign(info.style, {
      background: '#fff3cd', padding: '12px', borderRadius: '6px', marginBottom: '16px',
      border: '1px solid #ffc107', color: '#856404'
    });
    info.textContent = `A thread is currently open (${visibleCount}+ messages visible).`;
    
    const question = document.createElement('p');
    question.textContent = 'What would you like to export?';
    Object.assign(question.style, { margin: '16px 0', color: '#242424' });
    
    const buttonArea = document.createElement('div');
    Object.assign(buttonArea.style, {
      display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '20px'
    });
    
    const threadBtn = document.createElement('button');
    threadBtn.textContent = '💬 Export This Thread Only';
    Object.assign(threadBtn.style, {
      padding: '12px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#6264a7', color: '#fff', fontSize: '14px'
    });
    
    const fullChatBtn = document.createElement('button');
    fullChatBtn.textContent = '📋 Export Full Chat (close thread first)';
    Object.assign(fullChatBtn.style, {
      padding: '12px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#f0f0f0', fontSize: '14px', color: '#242424'
    });
    
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    Object.assign(cancelBtn.style, {
      padding: '10px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: 'transparent', color: '#666', fontSize: '14px'
    });
    
    buttonArea.appendChild(threadBtn);
    buttonArea.appendChild(fullChatBtn);
    buttonArea.appendChild(cancelBtn);
    
    modal.appendChild(title);
    modal.appendChild(info);
    modal.appendChild(question);
    modal.appendChild(buttonArea);
    
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    
    overlay.onclick = close;
    cancelBtn.onclick = close;
    
    // Export just the thread
    threadBtn.onclick = async () => {
      buttonArea.style.display = 'none';
      question.style.display = 'none';
      info.style.background = '#e3f2fd';
      info.style.border = '1px solid #2196f3';
      info.textContent = 'Scrolling through thread to capture all messages...';
      
      const threadMessages = await extractAllFromThread(rightRail);
      
      const md = generateMarkdown(threadMessages, chatTitle, true);
      try {
        await navigator.clipboard.writeText(md);
        info.style.background = '#d4edda';
        info.style.border = '1px solid #28a745';
        info.textContent = `✓ Copied ${threadMessages.length} thread messages to clipboard!`;
        setTimeout(close, 2000);
      } catch (e) {
        console.log('=== MARKDOWN OUTPUT ===\n', md);
        info.textContent = 'Clipboard blocked - check console (F12)';
      }
    };
    
    // Close thread and show full chat export UI
    fullChatBtn.onclick = async () => {
      // Close the thread panel
      const toggleBtn = document.querySelector('[data-tid="thread-list-pane-toggle-button"]');
      if (toggleBtn) {
        toggleBtn.click();
        await new Promise(r => setTimeout(r, 500));
      }
      // Remove current modal and show full chat UI
      modal.remove();
      showFullChatUI();
    };
    
  } else {
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    showFullChatUI();
  }
  
  function showFullChatUI() {
    // Clear modal content (using DOM methods to avoid Trusted Types violation)
    while (modal.firstChild) {
      modal.removeChild(modal.firstChild);
    }
    
    const title = document.createElement('h2');
    title.textContent = 'Export Teams Chat';
    Object.assign(title.style, { margin: '0 0 16px', color: '#242424' });
    
    const info = document.createElement('div');
    info.textContent = 'Chat: ' + chatTitle;
    Object.assign(info.style, {
      background: '#f5f5f5', padding: '12px', borderRadius: '6px', marginBottom: '16px',
      color: '#242424'
    });
    
    const label = document.createElement('label');
    label.textContent = 'Days to capture: ';
    Object.assign(label.style, { display: 'block', marginBottom: '8px', color: '#242424' });
    
    const input = document.createElement('input');
    input.type = 'number';
    input.value = '2';
    input.min = '1';
    input.max = '30';
    Object.assign(input.style, {
      width: '100%', padding: '8px', border: '1px solid #ddd',
      borderRadius: '6px', marginTop: '4px', boxSizing: 'border-box',
      color: '#242424', background: '#fff'
    });
    label.appendChild(input);
    
    // Thread expansion checkbox
    const threadLabel = document.createElement('label');
    Object.assign(threadLabel.style, {
      display: 'flex', alignItems: 'center', gap: '8px', marginTop: '12px', cursor: 'pointer',
      color: '#242424'
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
      cursor: 'pointer', background: '#f0f0f0', color: '#242424'
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
    
    if (!document.body.contains(modal)) {
      document.body.appendChild(modal);
    }
    
    cancelBtn.onclick = close;
    overlay.onclick = close;

    // Find a message's thread button by matching content
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
          let isThreadPreview = false, isSentToChannel = false;
          
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
            // Track "Also sent to channel" messages - these are duplicates of thread replies
            if (content.match(/^Also sent to channel\s*/i)) {
              isSentToChannel = true;
              content = content.replace(/^Also sent to channel\s*/i, '');
            }
            edited = !!item.querySelector('[id^="edited-"]');
            
            links = [...msg.querySelectorAll('a[href]')]
              .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
              .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
            
            reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
              .map(r => r.textContent?.trim()).filter(Boolean);
            
            // Check if this is a "Replied in thread" preview
            if (content.startsWith('Replied in thread:')) {
              isThreadPreview = true;
              content = content.replace(/^Replied in thread:\s*/, '');
            }
            
            // Check for thread replies
            const replySummary = item.querySelector('[data-tid="replies-summary-authors"]');
            if (replySummary && !isThreadPreview) {
              const summaryParent = replySummary.closest('[class*="repl"]') || replySummary.parentElement?.parentElement;
              const summaryText = summaryParent?.textContent || '';
              
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
                isSentToChannel,
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
            
            threadButton = findThreadButton(chatPane, m.contentSnippet, m.isoDate);
            
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
              
              const threadRail = document.querySelector('[data-tid="right-rail-message-pane-body"]');
              if (threadRail) {
                await new Promise(r => setTimeout(r, 500));
                
                // Find the scrollable container - walk up DOM
                let scrollContainer = threadRail;
                let el = threadRail;
                while (el && el !== document.body) {
                  if (el.scrollHeight > el.clientHeight + 10) {
                    scrollContainer = el;
                    break;
                  }
                  el = el.parentElement;
                }
                
                // Scroll through thread to load all replies (virtual scrolling)
                const threadMessages = new Map();
                const extractThreadMessages = () => {
                  threadRail.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
                    const msg = item.querySelector('[data-tid="chat-pane-message"]');
                    if (msg) {
                      const sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
                      const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
                      const isoDate = timeEl?.getAttribute('datetime') || '';
                      const content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
                      const edited = !!item.querySelector('[id^="edited-"]');
                      const links = [...msg.querySelectorAll('a[href]')]
                        .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
                        .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
                      const reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
                        .map(r => r.textContent?.trim()).filter(Boolean);
                      
                      if (content) {
                        const key = `${sender}-${isoDate}-${content.substring(0, 40)}`;
                        if (!threadMessages.has(key)) {
                          threadMessages.set(key, {
                            sender, isoDate, content, edited,
                            links: links.length ? links : null,
                            reactions: reactions.length ? reactions : null
                          });
                        }
                      }
                    }
                  });
                };
                
                // Scroll to BOTTOM first (newest messages), then scroll UP to get older
                scrollContainer.scrollTop = scrollContainer.scrollHeight;
                await new Promise(r => setTimeout(r, 300));
                extractThreadMessages();
                
                // Scroll UP through thread to load older messages
                let scrollAttempts = 0;
                const maxScrollAttempts = 50;
                while (scrollAttempts < maxScrollAttempts) {
                  const prevCount = threadMessages.size;
                  const prevScroll = scrollContainer.scrollTop;
                  
                  scrollContainer.scrollTop -= 400;
                  await new Promise(r => setTimeout(r, 200));
                  extractThreadMessages();
                  
                  // Stop if we've hit the top and no new messages
                  if (threadMessages.size === prevCount && (scrollContainer.scrollTop === prevScroll || scrollContainer.scrollTop === 0)) {
                    break;
                  }
                  scrollAttempts++;
                }
                
                // Sort by timestamp and convert to array
                const replies = [...threadMessages.values()].sort((a, b) => {
                  const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
                  const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
                  return da - db;
                });
                
                if (replies.length > 1) {
                  m.threadInfo.replies = replies.slice(1);
                } else if (replies.length === 1) {
                  m.threadInfo.replies = replies;
                }
                console.log(`Thread "${m.contentSnippet.substring(0, 30)}...": ${m.threadInfo.replies.length} replies (scrolled ${scrollAttempts}x)`);
              }
              
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

      // Filter and generate markdown
      const filtered = [...messages.values()]
        .filter(m => !m.isThreadPreview)
        // If expanding threads, skip "sent to channel" copies - they'll appear in the thread
        .filter(m => !expandThreads || !m.isSentToChannel)
        .filter(m => {
          if (m.isoDate) return new Date(m.isoDate) >= cutoff;
          return true;
        });

      progressBar.style.width = '95%';
      progressText.textContent = 'Generating markdown...';

      const md = generateMarkdown(filtered, chatTitle, false);

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
  }
})();
