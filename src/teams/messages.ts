/**
 * Message retrieval functionality.
 * Handles fetching and parsing message content.
 */

import type { Page } from 'playwright';
import type { TeamsMessage } from '../types/teams.js';

// Message container selectors
const MESSAGE_SELECTORS = [
  '[data-tid="chat-pane-message"]',
  '[data-tid="message-container"]',
  '.message-list-item',
  '[role="listitem"][data-testid*="message"]',
];

/**
 * Extracts messages from the current view (chat or channel).
 */
export async function getVisibleMessages(
  page: Page,
  maxMessages: number = 50
): Promise<TeamsMessage[]> {
  const messages: TeamsMessage[] = [];

  for (const selector of MESSAGE_SELECTORS) {
    const elements = await page.locator(selector).all();
    
    for (const element of elements.slice(0, maxMessages)) {
      try {
        const message = await parseMessageElement(element);
        if (message) {
          messages.push(message);
        }
      } catch {
        // Continue to next message
      }
    }

    if (messages.length > 0) break;
  }

  return messages;
}

/**
 * Parses a message DOM element into a TeamsMessage object.
 */
async function parseMessageElement(
  element: Awaited<ReturnType<Page['locator']>>
): Promise<TeamsMessage | null> {
  // Extract message ID
  const id = await element.getAttribute('data-tid') ??
             await element.getAttribute('id') ??
             `msg-${Date.now()}`;

  // Extract message content
  const contentSelectors = [
    '[data-tid="message-body"]',
    '.message-body',
    '[role="document"]',
    'p',
  ];

  let content = '';
  for (const selector of contentSelectors) {
    try {
      const contentEl = element.locator(selector).first();
      const text = await contentEl.textContent();
      if (text) {
        content = text.trim();
        break;
      }
    } catch {
      continue;
    }
  }

  if (!content) {
    const fullText = await element.textContent();
    content = fullText?.trim() ?? '';
  }

  if (!content) {
    return null;
  }

  // Extract sender
  const senderSelectors = [
    '[data-tid="message-author"]',
    '.message-author',
    '[data-tid="sender-name"]',
  ];

  let sender = 'Unknown';
  for (const selector of senderSelectors) {
    try {
      const senderEl = element.locator(selector).first();
      const text = await senderEl.textContent();
      if (text) {
        sender = text.trim();
        break;
      }
    } catch {
      continue;
    }
  }

  // Extract timestamp
  const timeSelectors = [
    '[data-tid="message-timestamp"]',
    'time',
    '[datetime]',
  ];

  let timestamp = new Date().toISOString();
  for (const selector of timeSelectors) {
    try {
      const timeEl = element.locator(selector).first();
      const datetime = await timeEl.getAttribute('datetime') ?? 
                       await timeEl.textContent();
      if (datetime) {
        timestamp = datetime.trim();
        break;
      }
    } catch {
      continue;
    }
  }

  return {
    id,
    content,
    sender,
    timestamp,
  };
}

/**
 * Scrolls to load more messages in the current view.
 */
export async function loadMoreMessages(page: Page): Promise<void> {
  // Find the scrollable container
  const scrollContainers = [
    '[data-tid="message-list"]',
    '.message-list',
    '[role="main"]',
  ];

  for (const selector of scrollContainers) {
    try {
      const container = page.locator(selector).first();
      await container.evaluate((el) => {
        el.scrollTop = 0; // Scroll to top to load older messages
      });
      await page.waitForTimeout(1000);
      return;
    } catch {
      continue;
    }
  }
}
