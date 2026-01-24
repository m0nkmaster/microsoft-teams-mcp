/**
 * Teams search functionality.
 * Implements both API interception (preferred) and DOM scraping (fallback).
 */

import type { Page, Locator } from 'playwright';
import type { TeamsSearchResult, TeamsSearchResultsWithPagination, SearchPaginationResult } from '../types/teams.js';
import { setupApiInterceptor, type SearchResultsWithPagination } from './api-interceptor.js';
import { stripHtml } from '../utils/parsers.js';

// Search box selectors for Teams v2 web app
const SEARCH_SELECTORS = [
  '[data-tid="searchInputField"]',
  '[data-tid="app-search-input"]',
  'input[data-tid*="search"]',
  'input[placeholder*="Search" i]',
  'input[aria-label*="Search" i]',
  'input[type="search"]',
  '[data-tid="search-box"]',
  '[role="search"] input',
];

// Clickable search trigger selectors
const SEARCH_TRIGGER_SELECTORS = [
  '[data-tid="search-button"]',
  '[data-tid="app-bar-search"]',
  'button[aria-label*="Search" i]',
  '[aria-label*="Search" i][role="button"]',
];

// DOM selectors from the Teams bookmarklet (proven to work)
// These target the chat/message view structure
const MESSAGE_SELECTORS = {
  item: '[data-tid="chat-pane-item"]',
  message: '[data-tid="chat-pane-message"]',
  controlMessage: '[data-tid="control-message-renderer"]',
  authorName: '[data-tid="message-author-name"]',
  timestamp: '[id^="timestamp-"]',
  time: 'time',
  content: '[id^="content-"]:not([id^="content-control"])',
  edited: '[id^="edited-"]',
  reactions: '[data-tid="diverse-reaction-pill-button"]',
};

// Search result specific selectors
const SEARCH_RESULT_SELECTORS = [
  '[data-tid*="search-result"]',
  '[data-tid*="message-result"]',
  '[data-tid*="searchResult"]',
  '[data-tid*="result-item"]',
  '[role="listitem"]:has([role="img"])',
  '[role="option"]:has([role="img"])',
];

export interface SearchOptions {
  /** Maximum results to return (for backward compat). Default: 25 */
  maxResults?: number;
  /** Timeout for waiting for results. Default: 10000 */
  waitMs?: number;
  /** Whether to use API interception (preferred). Default: true */
  useApiInterception?: boolean;
  /** Enable debug logging. Default: false */
  debug?: boolean;
  /** Starting offset for pagination (0-based). Default: 0 */
  from?: number;
  /** Page size. Default: 25 */
  size?: number;
}

const DEFAULT_OPTIONS: Required<SearchOptions> = {
  maxResults: 25,
  waitMs: 10000,
  useApiInterception: true,
  debug: false,
  from: 0,
  size: 25,
};

/**
 * Finds a working search input element on the page.
 */
async function findSearchInput(page: Page, debug = false): Promise<Locator | null> {
  for (const selector of SEARCH_SELECTORS) {
    try {
      const locator = page.locator(selector).first();
      const count = await locator.count();
      if (count > 0) {
        const isVisible = await locator.isVisible().catch(() => false);
        if (isVisible) {
          if (debug) console.log(`  [dom] Found search input: ${selector}`);
          return locator;
        }
      }
    } catch {
      continue;
    }
  }
  return null;
}

/**
 * Finds a clickable search trigger element.
 */
async function findSearchTrigger(page: Page, debug = false): Promise<Locator | null> {
  for (const selector of SEARCH_TRIGGER_SELECTORS) {
    try {
      const locator = page.locator(selector).first();
      const count = await locator.count();
      if (count > 0) {
        const isVisible = await locator.isVisible().catch(() => false);
        if (isVisible) {
          if (debug) console.log(`  [dom] Found search trigger: ${selector}`);
          return locator;
        }
      }
    } catch {
      continue;
    }
  }
  return null;
}

/**
 * Opens the search interface using various methods.
 */
async function openSearch(page: Page, debug = false): Promise<Locator> {
  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(2000);
  
  // Check if search input is already visible
  let searchInput = await findSearchInput(page, debug);
  if (searchInput) {
    return searchInput;
  }

  // Try keyboard shortcuts
  const isMac = process.platform === 'darwin';
  const shortcuts = [
    isMac ? 'Meta+e' : 'Control+e',
    isMac ? 'Meta+f' : 'Control+f',
    'F3',
  ];

  for (const shortcut of shortcuts) {
    if (debug) console.log(`  [dom] Trying shortcut: ${shortcut}`);
    await page.keyboard.press(shortcut);
    await page.waitForTimeout(1000);
    
    searchInput = await findSearchInput(page, debug);
    if (searchInput) {
      return searchInput;
    }
    
    await page.keyboard.press('Escape');
    await page.waitForTimeout(300);
  }

  // Try clicking search trigger buttons
  const searchTrigger = await findSearchTrigger(page, debug);
  if (searchTrigger) {
    if (debug) console.log('  [dom] Clicking search trigger');
    await searchTrigger.click();
    await page.waitForTimeout(1000);
    
    searchInput = await findSearchInput(page, debug);
    if (searchInput) {
      return searchInput;
    }
  }

  throw new Error(
    'Could not find search input. Teams UI may have changed. ' +
    'Run with debug:true to see what elements are available.'
  );
}

/**
 * Types a search query and submits it.
 */
async function typeSearchQuery(
  page: Page, 
  searchInput: Locator, 
  query: string,
  debug = false
): Promise<void> {
  if (debug) console.log(`  [dom] Typing query: "${query}"`);
  
  await searchInput.scrollIntoViewIfNeeded().catch(() => {});
  await page.waitForTimeout(300);
  await searchInput.waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
  
  // Try multiple interaction strategies
  let typed = false;
  
  // Strategy 1: Direct fill
  try {
    await searchInput.fill(query, { timeout: 5000 });
    typed = true;
    if (debug) console.log('  [dom] Used fill() strategy');
  } catch (e) {
    if (debug) console.log(`  [dom] fill() failed: ${e instanceof Error ? e.message : e}`);
  }
  
  // Strategy 2: Click then type
  if (!typed) {
    try {
      await searchInput.click({ timeout: 5000 });
      await page.waitForTimeout(200);
      await page.keyboard.type(query, { delay: 30 });
      typed = true;
      if (debug) console.log('  [dom] Used click+keyboard.type() strategy');
    } catch (e) {
      if (debug) console.log(`  [dom] click+type failed: ${e instanceof Error ? e.message : e}`);
    }
  }
  
  // Strategy 3: Focus via JavaScript
  if (!typed) {
    try {
      await searchInput.evaluate((el) => {
        (el as HTMLInputElement).focus();
        (el as HTMLInputElement).value = '';
      });
      await page.keyboard.type(query, { delay: 30 });
      typed = true;
      if (debug) console.log('  [dom] Used JS focus+type strategy');
    } catch (e) {
      if (debug) console.log(`  [dom] JS focus+type failed: ${e instanceof Error ? e.message : e}`);
    }
  }
  
  if (!typed) {
    throw new Error('Failed to type into search input using all strategies');
  }
  
  await page.waitForTimeout(500);
  await page.keyboard.press('Enter');
  
  if (debug) console.log('  [dom] Query submitted');
}

// stripHtml imported from ../utils/parsers.js

/**
 * Extracts search results from the DOM using bookmarklet-inspired selectors.
 * Fallback when API interception doesn't capture results.
 */
async function extractResultsFromDom(
  page: Page, 
  maxResults: number,
  debug = false
): Promise<TeamsSearchResult[]> {
  const results: TeamsSearchResult[] = [];
  const seenContent = new Set<string>();
  
  if (debug) console.log('  [dom] Extracting results from DOM...');

  // Strategy 1: Look for chat-pane-item elements (bookmarklet pattern)
  const items = await page.locator(MESSAGE_SELECTORS.item).all();
  
  if (debug) console.log(`  [dom] Found ${items.length} chat-pane-item elements`);
  
  for (const item of items) {
    if (results.length >= maxResults) break;
    
    try {
      // Skip control/system messages
      const isControl = await item.locator(MESSAGE_SELECTORS.controlMessage).count() > 0;
      if (isControl) continue;
      
      // Check it's a message
      const hasMessage = await item.locator(MESSAGE_SELECTORS.message).count() > 0;
      if (!hasMessage) continue;
      
      // Extract sender
      const sender = await item.locator(MESSAGE_SELECTORS.authorName).textContent().catch(() => null);
      
      // Extract timestamp (try timestamp id first, then time element)
      let timestamp: string | undefined;
      const timestampEl = item.locator(MESSAGE_SELECTORS.timestamp).first();
      if (await timestampEl.count() > 0) {
        timestamp = await timestampEl.getAttribute('datetime').catch(() => null) || undefined;
      }
      if (!timestamp) {
        const timeEl = item.locator(MESSAGE_SELECTORS.time).first();
        if (await timeEl.count() > 0) {
          timestamp = await timeEl.getAttribute('datetime').catch(() => null) || undefined;
        }
      }
      
      // Extract content
      const contentEl = item.locator(MESSAGE_SELECTORS.content).first();
      let content = '';
      if (await contentEl.count() > 0) {
        const html = await contentEl.innerHTML().catch(() => null);
        content = html ? stripHtml(html) : '';
      }
      
      // Skip empty or too short content
      if (content.length < 5) continue;
      
      // Deduplicate
      const key = content.substring(0, 60).toLowerCase();
      if (seenContent.has(key)) continue;
      seenContent.add(key);
      
      results.push({
        id: `dom-${results.length}`,
        type: 'message',
        content,
        sender: sender?.trim() || undefined,
        timestamp,
      });
    } catch {
      continue;
    }
  }
  
  // Strategy 2: Try search result specific selectors
  if (results.length < maxResults) {
    for (const selector of SEARCH_RESULT_SELECTORS) {
      const elements = await page.locator(selector).all();
      
      if (debug && elements.length > 0) {
        console.log(`  [dom] Found ${elements.length} elements with: ${selector}`);
      }
      
      for (const element of elements) {
        if (results.length >= maxResults) break;
        
        try {
          const text = await element.textContent();
          if (!text || text.length < 20) continue;
          
          const cleaned = stripHtml(text).replace(/\s+/g, ' ').trim();
          if (cleaned.length < 15) continue;
          
          // Deduplicate
          const key = cleaned.substring(0, 60).toLowerCase();
          if (seenContent.has(key)) continue;
          seenContent.add(key);
          
          // Try to parse sender and timestamp from text
          const parsed = parseResultText(cleaned);
          
          results.push({
            id: `result-${results.length}`,
            type: 'message',
            content: parsed.content,
            sender: parsed.sender,
            timestamp: parsed.timestamp,
          });
        } catch {
          continue;
        }
      }
      
      if (results.length >= maxResults) break;
    }
  }
  
  if (debug) console.log(`  [dom] Extracted ${results.length} results from DOM`);
  
  return results;
}

interface ParsedResult {
  sender?: string;
  timestamp?: string;
  content: string;
}

/**
 * Parses a result string to extract sender, timestamp, and content.
 */
function parseResultText(text: string): ParsedResult {
  let remaining = text;
  
  // Pattern: "Lastname, Firstname" at the start
  const senderMatch = remaining.match(/^([A-Z][a-zA-Z'-]+,\s*[A-Z][a-zA-Z'-]+)/);
  let sender: string | undefined;
  
  if (senderMatch) {
    sender = senderMatch[1];
    remaining = remaining.slice(sender.length).trim();
  }
  
  // Pattern: date/time
  const timePatterns = [
    /^(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2})/,
    /^(Yesterday\s+\d{1,2}:\d{2})/i,
    /^(Today\s+\d{1,2}:\d{2})/i,
    /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2}:\d{2})/i,
  ];
  
  let timestamp: string | undefined;
  for (const pattern of timePatterns) {
    const match = remaining.match(pattern);
    if (match) {
      timestamp = match[0];
      remaining = remaining.slice(timestamp.length).trim();
      break;
    }
  }
  
  return { sender, timestamp, content: remaining };
}

/**
 * Main search function.
 * Searches Teams for messages matching the query.
 * 
 * Prefers API interception for structured results, falls back to DOM scraping.
 */
export async function searchTeams(
  page: Page,
  query: string,
  options: SearchOptions = {}
): Promise<TeamsSearchResult[]> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const debug = opts.debug;
  
  // Set up API interception before triggering search
  const interceptor = opts.useApiInterception ? setupApiInterceptor(page, debug) : null;
  
  try {
    // Open search and type query
    const searchInput = await openSearch(page, debug);
    await typeSearchQuery(page, searchInput, query, debug);
    
    // Wait for results - try API first, then DOM
    if (interceptor) {
      if (debug) console.log('  [api] Waiting for API results...');
      
      const apiResults = await interceptor.waitForSearchResults(opts.waitMs);
      
      if (apiResults.length > 0) {
        if (debug) console.log(`  [api] Got ${apiResults.length} results from API`);
        return apiResults.slice(0, opts.maxResults);
      }
      
      if (debug) console.log('  [api] No API results, falling back to DOM');
    }
    
    // Wait a bit for DOM to render
    await page.waitForTimeout(2000);
    
    // Fall back to DOM extraction
    return extractResultsFromDom(page, opts.maxResults, debug);
    
  } finally {
    // Clean up interceptor
    interceptor?.stop();
  }
}

/**
 * Main search function with pagination metadata.
 * Searches Teams for messages matching the query and returns pagination info.
 * 
 * The Substrate v2 query API uses from/size pagination:
 * - from: Starting offset (0, 25, 50, 75, 100...)
 * - size: Page size (default 25)
 */
export async function searchTeamsWithPagination(
  page: Page,
  query: string,
  options: SearchOptions = {}
): Promise<TeamsSearchResultsWithPagination> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const debug = opts.debug;
  
  const defaultPagination: SearchPaginationResult = {
    returned: 0,
    from: opts.from,
    size: opts.size,
    hasMore: false,
  };
  
  // Set up API interception before triggering search
  const interceptor = opts.useApiInterception ? setupApiInterceptor(page, debug) : null;
  
  try {
    // Open search and type query
    const searchInput = await openSearch(page, debug);
    await typeSearchQuery(page, searchInput, query, debug);
    
    // Wait for results - try API first, then DOM
    if (interceptor) {
      if (debug) console.log('  [api] Waiting for API results with pagination...');
      
      const apiResults = await interceptor.waitForSearchResultsWithPagination(opts.waitMs);
      
      if (apiResults.results.length > 0) {
        if (debug) {
          console.log(`  [api] Got ${apiResults.results.length} results from API`);
          console.log(`  [api] Pagination: from=${apiResults.pagination.from}, size=${apiResults.pagination.size}, hasMore=${apiResults.pagination.hasMore}`);
        }
        
        return {
          results: apiResults.results.slice(0, opts.maxResults),
          pagination: {
            returned: Math.min(apiResults.results.length, opts.maxResults),
            from: apiResults.pagination.from,
            size: apiResults.pagination.size,
            total: apiResults.pagination.total,
            hasMore: apiResults.pagination.hasMore,
          },
        };
      }
      
      if (debug) console.log('  [api] No API results, falling back to DOM');
    }
    
    // Wait a bit for DOM to render
    await page.waitForTimeout(2000);
    
    // Fall back to DOM extraction
    const domResults = await extractResultsFromDom(page, opts.maxResults, debug);
    
    return {
      results: domResults,
      pagination: {
        ...defaultPagination,
        returned: domResults.length,
        // For DOM extraction, we don't know the total, assume more if we hit maxResults
        hasMore: domResults.length >= opts.maxResults,
      },
    };
    
  } finally {
    // Clean up interceptor
    interceptor?.stop();
  }
}

/**
 * Filters messages in the current view (channel/chat).
 */
export async function filterCurrentMessages(
  page: Page,
  query: string
): Promise<TeamsSearchResult[]> {
  return searchTeams(page, query);
}
