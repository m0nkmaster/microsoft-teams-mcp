/**
 * Direct API client for Teams/Substrate search.
 * 
 * Extracts auth tokens from browser session state and makes
 * direct HTTP calls without needing an active browser.
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import type { TeamsSearchResult, SearchPaginationResult } from '../types/teams.js';
import {
  stripHtml,
  buildMessageLink,
  parseJwtProfile,
  parsePeopleResults,
  parseSearchResults,
  type PersonSearchResult,
  type UserProfile,
} from '../utils/parsers.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const SESSION_STATE_PATH = path.join(PROJECT_ROOT, 'session-state.json');
const TOKEN_CACHE_PATH = path.join(PROJECT_ROOT, 'token-cache.json');

interface TokenCache {
  substrateToken: string;
  substrateTokenExpiry: number;
  extractedAt: number;
}

interface TeamsTokenInfo {
  token: string;
  expiry: Date;
  userMri: string;  // User's Teams MRI (8:orgid:guid)
}

interface MessageAuthInfo {
  skypeToken: string;
  authToken: string;
  userMri: string;
}

interface DirectSearchResult {
  results: TeamsSearchResult[];
  pagination: SearchPaginationResult;
}

// Re-export UserProfile from parsers for backward compatibility
export type { UserProfile } from '../utils/parsers.js';

/**
 * Gets the current user's profile from cached JWT tokens.
 * 
 * Extracts user info from MSAL tokens stored in session state.
 * No API call needed - just parses existing tokens.
 */
export function getMe(): UserProfile | null {
  if (!fs.existsSync(SESSION_STATE_PATH)) {
    return null;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    const teamsOrigin = state.origins?.find((o: { origin: string }) => 
      o.origin === 'https://teams.microsoft.com'
    );

    if (!teamsOrigin) return null;

    // Look through localStorage for any JWT with user info
    for (const item of teamsOrigin.localStorage) {
      try {
        const val = JSON.parse(item.value);
        
        if (!val.secret || typeof val.secret !== 'string') continue;
        if (!val.secret.startsWith('ey')) continue;
        
        const parts = val.secret.split('.');
        if (parts.length !== 3) continue;
        
        const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
        
        // Use shared parsing function
        const profile = parseJwtProfile(payload);
        if (profile) {
          return profile;
        }
      } catch {
        continue;
      }
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Extracts the Substrate search token from session state.
 */
export function extractSubstrateToken(): { token: string; expiry: Date } | null {
  if (!fs.existsSync(SESSION_STATE_PATH)) {
    return null;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    const teamsOrigin = state.origins?.find((o: { origin: string }) => 
      o.origin === 'https://teams.microsoft.com'
    );

    if (!teamsOrigin) return null;

    for (const item of teamsOrigin.localStorage) {
      try {
        const val = JSON.parse(item.value);
        if (val.target?.includes('substrate.office.com/search/SubstrateSearch')) {
          const token = val.secret;
          
          // Parse JWT to get expiry
          const parts = token.split('.');
          if (parts.length === 3) {
            const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
            const expiry = new Date(payload.exp * 1000);
            return { token, expiry };
          }
        }
      } catch {
        continue;
      }
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Gets a valid token, either from cache or by extracting from session.
 */
export function getValidToken(): string | null {
  // Try cache first
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    try {
      const cache: TokenCache = JSON.parse(fs.readFileSync(TOKEN_CACHE_PATH, 'utf8'));
      if (cache.substrateTokenExpiry > Date.now()) {
        return cache.substrateToken;
      }
    } catch {
      // Cache invalid, continue to extraction
    }
  }

  // Extract from session state
  const extracted = extractSubstrateToken();
  if (!extracted) return null;

  // Check if not expired
  if (extracted.expiry.getTime() <= Date.now()) {
    return null;
  }

  // Cache the token
  const cache: TokenCache = {
    substrateToken: extracted.token,
    substrateTokenExpiry: extracted.expiry.getTime(),
    extractedAt: Date.now(),
  };
  fs.writeFileSync(TOKEN_CACHE_PATH, JSON.stringify(cache, null, 2));

  return extracted.token;
}

/**
 * Clears the token cache (forces re-extraction on next call).
 */
export function clearTokenCache(): void {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
}

/**
 * Checks if we have a valid token for direct API calls.
 */
export function hasValidToken(): boolean {
  return getValidToken() !== null;
}

/**
 * Gets token status for diagnostics.
 */
export function getTokenStatus(): {
  hasToken: boolean;
  expiresAt?: string;
  minutesRemaining?: number;
} {
  const extracted = extractSubstrateToken();
  if (!extracted) {
    return { hasToken: false };
  }

  const now = Date.now();
  const expiryMs = extracted.expiry.getTime();
  
  return {
    hasToken: expiryMs > now,
    expiresAt: extracted.expiry.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - now) / 1000 / 60)),
  };
}

/**
 * Makes a direct search API call to Substrate.
 */
export async function directSearch(
  query: string,
  options: { from?: number; size?: number; maxResults?: number } = {}
): Promise<DirectSearchResult> {
  const token = getValidToken();
  if (!token) {
    throw new Error('No valid token available. Browser login required.');
  }

  const from = options.from ?? 0;
  const size = options.size ?? 25;

  // Generate unique IDs for this request
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    entityRequests: [{
      entityType: 'Message',
      contentSources: ['Teams'],
      propertySet: 'Optimized',
      fields: [
        'Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String',
      ],
      query: {
        queryString: `${query} AND NOT (isClientSoftDeleted:TRUE)`,
        displayQueryString: query,
      },
      from,
      size,
      topResultsCount: 5,
    }],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ['Suggestion'],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Messages' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
      Name: 'powerbar',
    },
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };

  const response = await fetch('https://substrate.office.com/searchservice/api/v2/query', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Origin': 'https://teams.microsoft.com',
      'Referer': 'https://teams.microsoft.com/',
    },
    body: JSON.stringify(body),
  });

  if (response.status === 401) {
    clearTokenCache();
    throw new Error('Token expired or invalid. Browser login required.');
  }

  if (!response.ok) {
    throw new Error(`API error: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  
  // Use shared parsing function
  const { results, total } = parseSearchResults(
    data.EntitySets as unknown[] | undefined,
    from,
    size
  );

  const maxResults = options.maxResults ?? size;
  const limitedResults = results.slice(0, maxResults);

  return {
    results: limitedResults,
    pagination: {
      from,
      size,
      returned: limitedResults.length,
      total,
      hasMore: total !== undefined 
        ? from + results.length < total 
        : results.length >= size,
    },
  };
}

/**
 * Extracts the Teams API token and user MRI from session state.
 * This is different from the Substrate token - it's used for chat APIs.
 * 
 * The chat API requires a token with audience:
 * - https://chatsvcagg.teams.microsoft.com (preferred)
 * - https://api.spaces.skype.com (fallback)
 */
export function extractTeamsToken(): TeamsTokenInfo | null {
  if (!fs.existsSync(SESSION_STATE_PATH)) {
    return null;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    const teamsOrigin = state.origins?.find((o: { origin: string }) => 
      o.origin === 'https://teams.microsoft.com'
    );

    if (!teamsOrigin) return null;

    let chatToken: string | null = null;
    let chatTokenExpiry: Date | null = null;
    let skypeToken: string | null = null;
    let skypeTokenExpiry: Date | null = null;
    let userMri: string | null = null;

    for (const item of teamsOrigin.localStorage) {
      try {
        const val = JSON.parse(item.value);
        
        if (!val.target || !val.secret) continue;
        
        const secret = val.secret;
        if (typeof secret !== 'string' || !secret.startsWith('ey')) continue;
        
        const parts = secret.split('.');
        if (parts.length !== 3) continue;
        
        const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
        const tokenExpiry = new Date(payload.exp * 1000);
        
        // Extract user MRI from any token
        if (payload.oid && !userMri) {
          userMri = `8:orgid:${payload.oid}`;
        }
        
        // Prefer chatsvcagg.teams.microsoft.com token
        if (val.target.includes('chatsvcagg.teams.microsoft.com')) {
          if (!chatTokenExpiry || tokenExpiry > chatTokenExpiry) {
            chatToken = secret;
            chatTokenExpiry = tokenExpiry;
          }
        }
        
        // Fallback to api.spaces.skype.com token
        if (val.target.includes('api.spaces.skype.com')) {
          if (!skypeTokenExpiry || tokenExpiry > skypeTokenExpiry) {
            skypeToken = secret;
            skypeTokenExpiry = tokenExpiry;
          }
        }
      } catch {
        continue;
      }
    }

    // If we still don't have userMri, try to get it from the Substrate token
    if (!userMri) {
      const substrateInfo = extractSubstrateToken();
      if (substrateInfo) {
        try {
          const parts = substrateInfo.token.split('.');
          if (parts.length === 3) {
            const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
            if (payload.oid) {
              userMri = `8:orgid:${payload.oid}`;
            }
          }
        } catch {
          // ignore
        }
      }
    }

    // Prefer chatsvc token, fallback to skype token
    const token = chatToken || skypeToken;
    const expiry = chatToken ? chatTokenExpiry : skypeTokenExpiry;

    if (token && expiry && userMri && expiry.getTime() > Date.now()) {
      return { token, expiry, userMri };
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Extracts authentication info needed for sending messages.
 * Uses cookies (skypetoken_asm) which are required for the chatsvc API.
 */
export function extractMessageAuth(): MessageAuthInfo | null {
  if (!fs.existsSync(SESSION_STATE_PATH)) {
    return null;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    
    let skypeToken: string | null = null;
    let authToken: string | null = null;
    let userMri: string | null = null;

    // Extract tokens from cookies
    for (const cookie of state.cookies || []) {
      if (cookie.name === 'skypetoken_asm' && cookie.domain?.includes('teams.microsoft.com')) {
        skypeToken = cookie.value;
      }
      if (cookie.name === 'authtoken' && cookie.domain?.includes('teams.microsoft.com')) {
        // Decode the URL-encoded cookie value
        authToken = decodeURIComponent(cookie.value);
        if (authToken.startsWith('Bearer=')) {
          authToken = authToken.substring(7); // Remove "Bearer=" prefix
        }
      }
    }

    // Get userMri from token payload or localStorage
    if (skypeToken) {
      try {
        const parts = skypeToken.split('.');
        if (parts.length >= 2) {
          const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
          if (payload.skypeid) {
            userMri = payload.skypeid;
          }
        }
      } catch {
        // Not a JWT format, that's fine
      }
    }

    // Fallback to extracting userMri from authToken
    if (!userMri && authToken) {
      try {
        const parts = authToken.split('.');
        if (parts.length === 3) {
          const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
          if (payload.oid) {
            userMri = `8:orgid:${payload.oid}`;
          }
        }
      } catch {
        // ignore
      }
    }

    if (skypeToken && authToken && userMri) {
      return { skypeToken, authToken, userMri };
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Gets user's display name from session state.
 */
export function getUserDisplayName(): string | null {
  if (!fs.existsSync(SESSION_STATE_PATH)) {
    return null;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    const teamsOrigin = state.origins?.find((o: { origin: string }) => 
      o.origin === 'https://teams.microsoft.com'
    );

    if (!teamsOrigin) return null;

    for (const item of teamsOrigin.localStorage) {
      try {
        // Look for user profile data
        if (item.value?.includes('displayName') || item.value?.includes('givenName')) {
          const val = JSON.parse(item.value);
          if (val.displayName) return val.displayName;
          if (val.name?.displayName) return val.name.displayName;
        }
      } catch {
        continue;
      }
    }

    // Try to get from token
    const teamsToken = extractTeamsToken();
    if (teamsToken) {
      try {
        const parts = teamsToken.token.split('.');
        if (parts.length === 3) {
          const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
          if (payload.name) return payload.name;
        }
      } catch {
        // ignore
      }
    }
  } catch {
    return null;
  }

  return null;
}

export interface SendMessageResult {
  success: boolean;
  messageId?: string;
  timestamp?: number;
  error?: string;
}

// Re-export PersonSearchResult from parsers for backward compatibility
export type { PersonSearchResult } from '../utils/parsers.js';

/** Search people results with count. */
export interface PeopleSearchResults {
  results: PersonSearchResult[];
  returned: number;
}

/** Frequent contacts result. */
export interface FrequentContactsResult {
  contacts: PersonSearchResult[];
  returned: number;
}

/** A favourite/pinned conversation item. */
export interface FavoriteItem {
  conversationId: string;
  createdTime?: number;
  lastUpdatedTime?: number;
}

/** Response from getting favorites. */
export interface FavoritesResult {
  success: boolean;
  favorites: FavoriteItem[];
  folderHierarchyVersion?: number;
  folderId?: string;
  error?: string;
}

/** Result of modifying favorites. */
export interface FavoriteModifyResult {
  success: boolean;
  error?: string;
}

/** Result of saving/unsaving a message. */
export interface SaveMessageResult {
  success: boolean;
  conversationId?: string;
  messageId?: string;
  saved?: boolean;
  error?: string;
}

/**
 * Sends a message to a Teams conversation.
 * 
 * Uses the skypetoken_asm cookie for authentication, which is required
 * by the Teams chatsvc API.
 * 
 * @param conversationId - The conversation ID (e.g., "48:notes" for self-chat)
 * @param content - Message content (HTML supported)
 * @param region - Region for the API (default: "amer")
 */
export async function sendMessage(
  conversationId: string,
  content: string,
  region: string = 'amer'
): Promise<SendMessageResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, error: 'No valid authentication. Browser login required.' };
  }

  const displayName = getUserDisplayName() || 'User';

  // Generate unique message ID
  const clientMessageId = Date.now().toString();
  const now = new Date().toISOString();

  // Wrap content in paragraph if not already HTML
  const htmlContent = content.startsWith('<') ? content : `<p>${content}</p>`;

  const body = {
    content: htmlContent,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    imdisplayname: displayName,
    clientmessageid: clientMessageId,
  };

  // Use the Teams messaging API with skypetoken authentication
  const url = `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
        'X-Ms-Client-Version': '1415/1.0.0.2025010401',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    const data = await response.json();

    return {
      success: true,
      messageId: clientMessageId,
      timestamp: data.OriginalArrivalTime,
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

/**
 * Sends a message to your own notes/self-chat.
 */
export async function sendNoteToSelf(content: string): Promise<SendMessageResult> {
  return sendMessage('48:notes', content);
}

/**
 * Searches for people by name or email using the Substrate suggestions API.
 * 
 * Uses the same auth token as message search.
 * 
 * @param query - Search term (name, email, or partial match)
 * @param limit - Maximum number of results (default: 10)
 */
export async function searchPeople(
  query: string,
  limit: number = 10
): Promise<PeopleSearchResults> {
  const token = getValidToken();
  if (!token) {
    throw new Error('No valid token available. Browser login required.');
  }

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI', 
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
  };

  const response = await fetch(
    'https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify(body),
    }
  );

  if (response.status === 401) {
    clearTokenCache();
    throw new Error('Token expired or invalid. Browser login required.');
  }

  if (!response.ok) {
    throw new Error(`API error: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  
  // Use shared parsing function
  const results = parsePeopleResults(data.Groups as unknown[] | undefined);

  return {
    results,
    returned: results.length,
  };
}

/**
 * Gets the user's frequently contacted people.
 * 
 * Uses the peoplecache scenario which returns contacts ranked by
 * interaction frequency. Useful for resolving ambiguous names
 * (e.g., "Rob" → "Rob Smith <rob.smith@company.com>").
 * 
 * @param limit - Maximum number of contacts to return (default: 50)
 */
export async function getFrequentContacts(
  limit: number = 50
): Promise<FrequentContactsResult> {
  const token = getValidToken();
  if (!token) {
    throw new Error('No valid token available. Browser login required.');
  }

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: '',
        DisplayQueryString: '',
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
  };

  const response = await fetch(
    'https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify(body),
    }
  );

  if (response.status === 401) {
    clearTokenCache();
    throw new Error('Token expired or invalid. Browser login required.');
  }

  if (!response.ok) {
    throw new Error(`API error: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  
  // Use shared parsing function
  const contacts = parsePeopleResults(data.Groups as unknown[] | undefined);

  return {
    contacts,
    returned: contacts.length,
  };
}

/**
 * Gets the user's favourite/pinned conversations.
 * 
 * Uses the conversationFolders API with chatsvc authentication.
 * 
 * @param region - Region for the API (default: "amer")
 */
export async function getFavorites(region: string = 'amer'): Promise<FavoritesResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, favorites: [], error: 'No valid authentication. Browser login required.' };
  }

  const url = `https://teams.microsoft.com/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify({}),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        favorites: [], 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    const data = await response.json();
    
    // Find the Favorites folder
    const folders = data.conversationFolders as unknown[] | undefined;
    const favoritesFolder = folders?.find((f: unknown) => {
      const folder = f as Record<string, unknown>;
      return folder.folderType === 'Favorites';
    }) as Record<string, unknown> | undefined;

    if (!favoritesFolder) {
      return {
        success: true,
        favorites: [],
        folderHierarchyVersion: data.folderHierarchyVersion,
      };
    }

    const items = favoritesFolder.conversationFolderItems as unknown[] | undefined;
    const favorites: FavoriteItem[] = (items || []).map((item: unknown) => {
      const i = item as Record<string, unknown>;
      return {
        conversationId: i.conversationId as string,
        createdTime: i.createdTime as number | undefined,
        lastUpdatedTime: i.lastUpdatedTime as number | undefined,
      };
    });

    return {
      success: true,
      favorites,
      folderHierarchyVersion: data.folderHierarchyVersion,
      folderId: favoritesFolder.id as string,
    };
  } catch (err) {
    return {
      success: false,
      favorites: [],
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

/**
 * Adds a conversation to the user's favourites.
 * 
 * @param conversationId - The conversation ID to add
 * @param region - Region for the API (default: "amer")
 */
export async function addFavorite(
  conversationId: string,
  region: string = 'amer'
): Promise<FavoriteModifyResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, error: 'No valid authentication. Browser login required.' };
  }

  // First, get the current folder state to get the folderId and version
  const currentState = await getFavorites(region);
  if (!currentState.success) {
    return { success: false, error: currentState.error };
  }

  if (!currentState.folderId) {
    return { success: false, error: 'Could not find Favorites folder' };
  }

  const url = `https://teams.microsoft.com/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify({
        folderHierarchyVersion: currentState.folderHierarchyVersion,
        actions: [
          {
            action: 'AddItem',
            folderId: currentState.folderId,
            itemId: conversationId,
          },
        ],
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    return { success: true };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

/**
 * Removes a conversation from the user's favourites.
 * 
 * @param conversationId - The conversation ID to remove
 * @param region - Region for the API (default: "amer")
 */
export async function removeFavorite(
  conversationId: string,
  region: string = 'amer'
): Promise<FavoriteModifyResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, error: 'No valid authentication. Browser login required.' };
  }

  // First, get the current folder state to get the folderId and version
  const currentState = await getFavorites(region);
  if (!currentState.success) {
    return { success: false, error: currentState.error };
  }

  if (!currentState.folderId) {
    return { success: false, error: 'Could not find Favorites folder' };
  }

  const url = `https://teams.microsoft.com/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify({
        folderHierarchyVersion: currentState.folderHierarchyVersion,
        actions: [
          {
            action: 'RemoveItem',
            folderId: currentState.folderId,
            itemId: conversationId,
          },
        ],
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    return { success: true };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

/**
 * Saves (bookmarks) a message.
 * 
 * @param conversationId - The conversation ID containing the message
 * @param messageId - The message ID to save (numeric string)
 * @param region - Region for the API (default: "amer")
 */
export async function saveMessage(
  conversationId: string,
  messageId: string,
  region: string = 'amer'
): Promise<SaveMessageResult> {
  return setMessageSavedState(conversationId, messageId, true, region);
}

/**
 * Unsaves (removes bookmark from) a message.
 * 
 * @param conversationId - The conversation ID containing the message
 * @param messageId - The message ID to unsave (numeric string)
 * @param region - Region for the API (default: "amer")
 */
export async function unsaveMessage(
  conversationId: string,
  messageId: string,
  region: string = 'amer'
): Promise<SaveMessageResult> {
  return setMessageSavedState(conversationId, messageId, false, region);
}

/** A message from a thread/conversation. */
export interface ThreadMessage {
  id: string;                    // Message ID (numeric string timestamp)
  content: string;               // Message content (may contain HTML)
  contentType: string;           // e.g., "RichText/Html", "Text"
  sender: {
    mri: string;                 // Sender's MRI (8:orgid:guid)
    displayName?: string;        // Sender's display name
  };
  timestamp: string;             // ISO timestamp
  conversationId: string;        // Parent conversation ID
  clientMessageId?: string;      // Client-generated message ID
  isFromMe?: boolean;            // Whether this message is from the current user
  messageLink?: string;          // Direct link to open this message in Teams
}

/** Result of getting thread messages. */
export interface GetThreadResult {
  success: boolean;
  conversationId?: string;
  messages?: ThreadMessage[];
  error?: string;
}

/**
 * Gets messages from a Teams conversation/thread.
 * 
 * This retrieves messages from a conversation, which can be:
 * - A 1:1 or group chat
 * - A channel thread
 * - Self-notes (48:notes)
 * 
 * @param conversationId - The conversation ID (e.g., "19:abc@thread.tacv2")
 * @param options - Optional parameters for pagination
 * @param options.limit - Maximum number of messages to return (default: 50)
 * @param options.startTime - Only get messages after this timestamp (epoch ms)
 * @param region - Region for the API (default: "amer")
 */
export async function getThreadMessages(
  conversationId: string,
  options: { limit?: number; startTime?: number } = {},
  region: string = 'amer'
): Promise<GetThreadResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, error: 'No valid authentication. Browser login required.' };
  }

  const limit = options.limit ?? 50;
  
  // Build URL with query parameters
  let url = `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?view=msnp24Equivalent&pageSize=${limit}`;
  
  if (options.startTime) {
    url += `&startTime=${options.startTime}`;
  }

  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    const data = await response.json();
    
    // Parse messages from the response
    // The API returns { messages: [...] } array
    const rawMessages = data.messages as unknown[] | undefined;
    
    if (!Array.isArray(rawMessages)) {
      return {
        success: true,
        conversationId,
        messages: [],
      };
    }

    const messages: ThreadMessage[] = [];
    
    for (const raw of rawMessages) {
      const msg = raw as Record<string, unknown>;
      
      // Skip non-message types (e.g., typing indicators, control messages)
      const messageType = msg.messagetype as string;
      if (!messageType || messageType.startsWith('Control/') || messageType === 'ThreadActivity/AddMember') {
        continue;
      }
      
      const id = msg.id as string || msg.originalarrivaltime as string;
      if (!id) continue;
      
      const content = msg.content as string || '';
      const contentType = msg.messagetype as string || 'Text';
      
      // Parse sender info
      const fromMri = msg.from as string || '';
      const displayName = msg.imdisplayname as string || msg.displayName as string;
      
      const timestamp = msg.originalarrivaltime as string || 
                       msg.composetime as string || 
                       new Date(parseInt(id, 10)).toISOString();
      
      // Build message link - id is already the timestamp in milliseconds
      const messageLink = /^\d+$/.test(id) 
        ? buildMessageLink(conversationId, id)
        : undefined;
      
      messages.push({
        id,
        content: stripHtml(content),
        contentType,
        sender: {
          mri: fromMri,
          displayName,
        },
        timestamp,
        conversationId,
        clientMessageId: msg.clientmessageid as string,
        isFromMe: fromMri === auth.userMri,
        messageLink,
      });
    }

    // Sort by timestamp (oldest first)
    messages.sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());

    return {
      success: true,
      conversationId,
      messages,
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

// stripHtml is imported from ../utils/parsers.js

/**
 * Internal function to set the saved state of a message.
 */
async function setMessageSavedState(
  conversationId: string,
  messageId: string,
  saved: boolean,
  region: string
): Promise<SaveMessageResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return { success: false, error: 'No valid authentication. Browser login required.' };
  }

  const url = `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/rcmetadata/${messageId}`;

  try {
    const response = await fetch(url, {
      method: 'PUT',
      headers: {
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': 'https://teams.microsoft.com',
        'Referer': 'https://teams.microsoft.com/',
      },
      body: JSON.stringify({
        s: saved ? 1 : 0,
        mid: parseInt(messageId, 10),
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { 
        success: false, 
        error: `API error: ${response.status} - ${errorText}` 
      };
    }

    return {
      success: true,
      conversationId,
      messageId,
      saved,
    };
  } catch (err) {
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

// parsePersonSuggestion is imported from ../utils/parsers.js

// Re-export buildMessageLink for backward compatibility
export { buildMessageLink };

// extractMessageTimestamp is imported from ../utils/parsers.js

// parseV2Result is imported from ../utils/parsers.js
