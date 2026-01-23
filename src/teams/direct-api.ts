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

export interface UserProfile {
  id: string;           // Azure AD object ID (oid)
  mri: string;          // Teams MRI (8:orgid:guid)
  email: string;        // User principal name / email
  displayName: string;  // Full display name
  givenName?: string;   // First name
  surname?: string;     // Last name
  tenantId?: string;    // Azure tenant ID
}

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
        
        // Look for tokens with user info
        if (payload.oid && payload.name) {
          const profile: UserProfile = {
            id: payload.oid,
            mri: `8:orgid:${payload.oid}`,
            email: payload.upn || payload.preferred_username || payload.email || '',
            displayName: payload.name,
            tenantId: payload.tid,
          };
          
          // Try to extract given name and surname
          if (payload.given_name) {
            profile.givenName = payload.given_name;
          }
          if (payload.family_name) {
            profile.surname = payload.family_name;
          }
          
          // If no given/family name, try to parse from displayName
          if (!profile.givenName && profile.displayName.includes(',')) {
            // Format: "Surname, GivenName"
            const parts = profile.displayName.split(',').map(s => s.trim());
            if (parts.length === 2) {
              profile.surname = parts[0];
              profile.givenName = parts[1];
            }
          } else if (!profile.givenName && profile.displayName.includes(' ')) {
            // Format: "GivenName Surname"
            const parts = profile.displayName.split(' ');
            profile.givenName = parts[0];
            profile.surname = parts.slice(1).join(' ');
          }
          
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
  
  // Parse results
  const results: TeamsSearchResult[] = [];
  let total: number | undefined;

  const entitySets = data.EntitySets as unknown[] | undefined;
  if (Array.isArray(entitySets)) {
    for (const entitySet of entitySets) {
      const es = entitySet as Record<string, unknown>;
      const resultSets = es.ResultSets as unknown[] | undefined;
      
      if (Array.isArray(resultSets)) {
        for (const resultSet of resultSets) {
          const rs = resultSet as Record<string, unknown>;
          
          // Try to get total
          const rsTotal = rs.Total ?? rs.TotalCount ?? rs.TotalEstimate;
          if (typeof rsTotal === 'number') {
            total = rsTotal;
          }
          
          const items = rs.Results as unknown[] | undefined;
          if (Array.isArray(items)) {
            for (const item of items) {
              const parsed = parseV2Result(item as Record<string, unknown>);
              if (parsed) results.push(parsed);
            }
          }
        }
      }
    }
  }

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
 * Parses a v2 query result item.
 */
function parseV2Result(item: Record<string, unknown>): TeamsSearchResult | null {
  const content = item.HitHighlightedSummary as string || 
                  item.Summary as string || 
                  '';
  
  if (content.length < 5) return null;

  const id = item.Id as string || 
             item.ReferenceId as string || 
             `v2-${Date.now()}`;

  // Strip HTML from content
  const cleanContent = content
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();

  const source = item.Source as Record<string, unknown> | undefined;

  // Extract conversationId from extension fields
  // The API returns it as Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String
  let conversationId: string | undefined;
  if (source) {
    // Find the first valid string value from potential field names
    conversationId = [
      source['Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String'],
      source['SkypeGroupId'],
      source['ConversationId'],
      source['ThreadId'],
    ].find((id): id is string => typeof id === 'string' && id.length > 0);
  }

  return {
    id,
    type: 'message',
    content: cleanContent,
    sender: source?.From as string || source?.Sender as string,
    timestamp: source?.ReceivedTime as string || source?.CreatedDateTime as string,
    channelName: source?.ChannelName as string || source?.Topic as string,
    teamName: source?.TeamName as string || source?.GroupName as string,
    conversationId,
    messageId: item.ReferenceId as string,
  };
}
