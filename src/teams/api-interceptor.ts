/**
 * Network request interception for Teams API calls.
 * Captures search requests and responses for direct API usage.
 */

import type { Page, Response } from 'playwright';
import type { TeamsSearchResult } from '../types/teams.js';
import { stripHtml, parseV2Result } from '../utils/parsers.js';

// Patterns that indicate search-related API calls
const SEARCH_API_PATTERNS = [
  /substrate\.office\.com\/searchservice\/api\/v2\/query/i, // v2 query (full search)
  /substrate\.office\.com\/search\/api\/v1\/suggestions/i,  // v1 suggestions (typeahead)
  /substrate\.office\.com\/search/i,                        // Other substrate search
  /\/api\/csa\/.*\/containers\/.*\/posts/i,                 // Channel posts
  /\/api\/chatsvc\/.*\/messages/i,                          // Chat messages
];

export interface CapturedSearchRequest {
  url: string;
  method: string;
  headers: Record<string, string>;
  body?: string;
}

export interface CapturedSearchResponse {
  url: string;
  status: number;
  body?: unknown;
}

/** Pagination metadata extracted from search results. */
export interface PaginationInfo {
  /** Number of results returned in this response. */
  returned: number;
  /** Starting offset used for this request. */
  from: number;
  /** Page size requested. */
  size: number;
  /** Total results available (if known from response). */
  total?: number;
  /** Whether more results are available. */
  hasMore: boolean;
}

/** Extended search results with pagination metadata. */
export interface SearchResultsWithPagination {
  results: TeamsSearchResult[];
  pagination: PaginationInfo;
}

export interface ApiInterceptor {
  responses: CapturedSearchResponse[];
  waitForSearchResults(timeoutMs?: number): Promise<TeamsSearchResult[]>;
  waitForSearchResultsWithPagination(timeoutMs?: number): Promise<SearchResultsWithPagination>;
  stop(): void;
}

function isSearchApiUrl(url: string): boolean {
  return SEARCH_API_PATTERNS.some(pattern => pattern.test(url));
}

// stripHtml imported from ../utils/parsers.js

/**
 * Parses Substrate search API response.
 * Handles both suggestions endpoint and full query endpoint.
 */
function parseSubstrateResponse(body: unknown): TeamsSearchResult[] {
  if (!body || typeof body !== 'object') return [];
  
  const results: TeamsSearchResult[] = [];
  const obj = body as Record<string, unknown>;
  
  // Try Groups[].Suggestions[] structure (suggestions endpoint)
  const groups = obj.Groups as unknown[] | undefined;
  if (Array.isArray(groups)) {
    for (const group of groups) {
      const g = group as Record<string, unknown>;
      const suggestions = g.Suggestions as unknown[] | undefined;
      const entityType = g.Type as string | undefined;
      
      if (Array.isArray(suggestions)) {
        for (const suggestion of suggestions) {
          const s = suggestion as Record<string, unknown>;
          
          // Skip people suggestions, we want messages
          if (entityType === 'People') continue;
          
          const result = parseSubstrateItem(s);
          if (result) results.push(result);
        }
      }
    }
  }
  
  // Try EntitySets[].ResultSets[].Results[] structure (v2 query endpoint)
  // Delegate to parsers.ts parseV2Result for consistency
  const entitySets = obj.EntitySets as unknown[] | undefined;
  if (Array.isArray(entitySets)) {
    for (const entitySet of entitySets) {
      const es = entitySet as Record<string, unknown>;
      const resultSets = es.ResultSets as unknown[] | undefined;
      
      if (Array.isArray(resultSets)) {
        for (const resultSet of resultSets) {
          const rs = resultSet as Record<string, unknown>;
          const items = rs.Results as unknown[] | undefined;
          
          if (Array.isArray(items)) {
            for (const item of items) {
              const result = parseV2Result(item as Record<string, unknown>);
              if (result) results.push(result);
            }
          }
        }
      }
    }
  }
  
  // Try value[] array (common Microsoft API pattern)
  const value = obj.value as unknown[] | undefined;
  if (Array.isArray(value)) {
    for (const item of value) {
      const result = parseSubstrateItem(item as Record<string, unknown>);
      if (result) results.push(result);
    }
  }
  
  return results;
}

/**
 * Extracts pagination metadata from v2 query response.
 */
function extractV2Pagination(body: unknown): { total?: number; returned: number } {
  if (!body || typeof body !== 'object') return { returned: 0 };
  
  const obj = body as Record<string, unknown>;
  let returned = 0;
  let total: number | undefined;
  
  const entitySets = obj.EntitySets as unknown[] | undefined;
  if (Array.isArray(entitySets)) {
    for (const entitySet of entitySets) {
      const es = entitySet as Record<string, unknown>;
      
      // Try to get total from EntitySet
      const esTotal = es.Total ?? es.TotalCount ?? es.Count;
      if (typeof esTotal === 'number') {
        total = esTotal;
      }
      
      const resultSets = es.ResultSets as unknown[] | undefined;
      if (Array.isArray(resultSets)) {
        for (const resultSet of resultSets) {
          const rs = resultSet as Record<string, unknown>;
          
          // Try to get total from ResultSet
          const rsTotal = rs.Total ?? rs.TotalCount ?? rs.TotalEstimate;
          if (typeof rsTotal === 'number' && total === undefined) {
            total = rsTotal;
          }
          
          const items = rs.Results as unknown[] | undefined;
          if (Array.isArray(items)) {
            returned += items.length;
          }
        }
      }
    }
  }
  
  return { total, returned };
}

/**
 * Parses a single Substrate search result item (suggestions endpoint).
 */
function parseSubstrateItem(item: Record<string, unknown>): TeamsSearchResult | null {
  // Try to extract content from various possible locations
  const content = 
    getStringField(item, 'Summary') ||
    getStringField(item, 'Body') ||
    getStringField(item, 'Content') ||
    getStringField(item, 'Snippet') ||
    getNestedContent(item, 'HitHighlightedSummary') ||
    getNestedContent(item, 'Preview') ||
    '';
  
  // Skip if no meaningful content
  if (content.length < 5) return null;
  
  const id = 
    getStringField(item, 'Id') ||
    getStringField(item, 'DocId') ||
    getStringField(item, 'ResourceId') ||
    `substrate-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  
  const sender = 
    getStringField(item, 'From') ||
    getStringField(item, 'Author') ||
    getStringField(item, 'DisplayName') ||
    getNestedName(item, 'Sender') ||
    getNestedName(item, 'From');
  
  const timestamp = 
    getStringField(item, 'ReceivedTime') ||
    getStringField(item, 'CreatedDateTime') ||
    getStringField(item, 'LastModifiedTime') ||
    getStringField(item, 'SentDateTime');
  
  const channelName = 
    getStringField(item, 'ChannelName') ||
    getStringField(item, 'Topic');
  
  const teamName = 
    getStringField(item, 'TeamName') ||
    getStringField(item, 'GroupName');
  
  return {
    id,
    type: 'message',
    content: stripHtml(content),
    sender,
    timestamp,
    channelName,
    teamName,
  };
}

// Note: parseV2QueryItem functionality is now delegated to parseV2Result in parsers.ts

/**
 * Parses Teams posts API response.
 */
function parsePostsResponse(body: unknown): TeamsSearchResult[] {
  if (!body || typeof body !== 'object') return [];
  
  const results: TeamsSearchResult[] = [];
  const obj = body as Record<string, unknown>;
  
  const posts = obj.posts as unknown[] | undefined;
  if (!Array.isArray(posts)) return [];
  
  for (const post of posts) {
    const p = post as Record<string, unknown>;
    const message = p.message as Record<string, unknown> | undefined;
    
    if (!message) continue;
    
    const content = getStringField(message, 'content') || '';
    if (content.length < 5) continue;
    
    const id = getStringField(p, 'id') || `post-${Date.now()}`;
    const sender = [
      getStringField(message, 'fromGivenNameInToken'),
      getStringField(message, 'fromFamilyNameInToken'),
    ].filter(Boolean).join(' ') || getStringField(message, 'imdisplayname');
    
    const timestamp = getStringField(p, 'latestMessageTime') || getStringField(p, 'originalArrivalTime');
    
    results.push({
      id,
      type: 'message',
      content: stripHtml(content),
      sender,
      timestamp,
      conversationId: getStringField(p, 'containerId'),
      messageId: id,
    });
  }
  
  return results;
}

function getStringField(obj: Record<string, unknown>, field: string): string | undefined {
  const value = obj[field];
  return typeof value === 'string' ? value : undefined;
}

function getNestedContent(obj: Record<string, unknown>, field: string): string | undefined {
  const value = obj[field];
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object') {
    const nested = value as Record<string, unknown>;
    return getStringField(nested, 'Content') || getStringField(nested, 'Text');
  }
  return undefined;
}

function getNestedName(obj: Record<string, unknown>, field: string): string | undefined {
  const value = obj[field];
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object') {
    const nested = value as Record<string, unknown>;
    return (
      getStringField(nested, 'DisplayName') ||
      getStringField(nested, 'Name') ||
      getStringField(nested, 'EmailAddress')
    );
  }
  return undefined;
}

/**
 * Parses a captured response based on its URL pattern.
 */
export function parseSearchResults(response: CapturedSearchResponse): TeamsSearchResult[] {
  if (!response.body || response.status >= 400) return [];
  
  const url = response.url.toLowerCase();
  
  if (url.includes('substrate.office.com')) {
    return parseSubstrateResponse(response.body);
  }
  
  if (url.includes('/posts') || url.includes('/csa/')) {
    return parsePostsResponse(response.body);
  }
  
  // Try generic parsing
  return parseSubstrateResponse(response.body);
}

/**
 * Parses a captured response and returns results with pagination info.
 */
export function parseSearchResultsWithPagination(
  response: CapturedSearchResponse,
  requestedFrom = 0,
  requestedSize = 25
): SearchResultsWithPagination {
  const results = parseSearchResults(response);
  
  // Extract pagination from v2 query response
  let pagination: PaginationInfo = {
    returned: results.length,
    from: requestedFrom,
    size: requestedSize,
    hasMore: results.length >= requestedSize,
  };
  
  if (response.url.includes('searchservice/api/v2/query')) {
    const v2Pagination = extractV2Pagination(response.body);
    pagination = {
      returned: v2Pagination.returned,
      from: requestedFrom,
      size: requestedSize,
      total: v2Pagination.total,
      hasMore: v2Pagination.total !== undefined 
        ? requestedFrom + v2Pagination.returned < v2Pagination.total
        : v2Pagination.returned >= requestedSize,
    };
  }
  
  return { results, pagination };
}

/**
 * Sets up request/response interception on a page.
 * Returns an interceptor with a promise-based wait for results.
 */
export function setupApiInterceptor(page: Page, debug = false): ApiInterceptor {
  const responses: CapturedSearchResponse[] = [];
  let resolveWait: ((results: TeamsSearchResult[]) => void) | null = null;
  let stopped = false;

  const responseHandler = async (response: Response) => {
    if (stopped) return;
    
    const url = response.url();
    if (!isSearchApiUrl(url)) return;
    
    if (debug) {
      console.log(`  [api] Captured response: ${url.substring(0, 80)}...`);
    }
    
    let body: unknown;
    try {
      const contentType = response.headers()['content-type'] || '';
      if (contentType.includes('application/json')) {
        body = await response.json();
      }
    } catch {
      // Response body may not be available
      return;
    }

    const captured: CapturedSearchResponse = {
      url,
      status: response.status(),
      body,
    };
    
    responses.push(captured);
    
    // Parse results and resolve if we got something
    const results = parseSearchResults(captured);
    if (results.length > 0 && resolveWait) {
      if (debug) {
        console.log(`  [api] Parsed ${results.length} results from ${url.substring(0, 60)}...`);
      }
      resolveWait(results);
      resolveWait = null;
    }
  };

  page.on('response', responseHandler);

  return {
    responses,
    
    waitForSearchResults(timeoutMs = 10000): Promise<TeamsSearchResult[]> {
      return new Promise((resolve) => {
        // Check if we already have results
        for (const resp of responses) {
          const results = parseSearchResults(resp);
          if (results.length > 0) {
            resolve(results);
            return;
          }
        }
        
        // Wait for new results
        resolveWait = resolve;
        
        // Timeout fallback
        setTimeout(() => {
          if (resolveWait) {
            resolveWait = null;
            
            // Try to get any results we have
            for (const resp of responses) {
              const results = parseSearchResults(resp);
              if (results.length > 0) {
                resolve(results);
                return;
              }
            }
            
            resolve([]);
          }
        }, timeoutMs);
      });
    },
    
    waitForSearchResultsWithPagination(timeoutMs = 10000): Promise<SearchResultsWithPagination> {
      return new Promise((resolve) => {
        const defaultPagination: PaginationInfo = {
          returned: 0,
          from: 0,
          size: 25,
          hasMore: false,
        };
        
        // Check if we already have results
        for (const resp of responses) {
          const parsed = parseSearchResultsWithPagination(resp);
          if (parsed.results.length > 0) {
            resolve(parsed);
            return;
          }
        }
        
        // Wait for new results
        const resolveWithPagination = (results: TeamsSearchResult[]) => {
          // Find the response that gave us these results
          for (const resp of responses) {
            const parsed = parseSearchResultsWithPagination(resp);
            if (parsed.results.length > 0) {
              resolve(parsed);
              return;
            }
          }
          resolve({ results, pagination: { ...defaultPagination, returned: results.length } });
        };
        
        resolveWait = resolveWithPagination;
        
        // Timeout fallback
        setTimeout(() => {
          if (resolveWait) {
            resolveWait = null;
            
            for (const resp of responses) {
              const parsed = parseSearchResultsWithPagination(resp);
              if (parsed.results.length > 0) {
                resolve(parsed);
                return;
              }
            }
            
            resolve({ results: [], pagination: defaultPagination });
          }
        }, timeoutMs);
      });
    },
    
    stop() {
      stopped = true;
      page.off('response', responseHandler);
    },
  };
}
