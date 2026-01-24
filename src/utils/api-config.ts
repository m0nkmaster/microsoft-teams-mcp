/**
 * API endpoint configuration and header utilities.
 * 
 * Centralises all API URLs and common request headers.
 */

/** Valid API regions. */
export const VALID_REGIONS = ['amer', 'emea', 'apac'] as const;
export type Region = typeof VALID_REGIONS[number];

/**
 * Validates a region string.
 * @throws Error if region is invalid.
 */
export function validateRegion(region: string): Region {
  if (!VALID_REGIONS.includes(region as Region)) {
    throw new Error(
      `Invalid region: ${region}. Valid regions: ${VALID_REGIONS.join(', ')}`
    );
  }
  return region as Region;
}

/**
 * Attempts to parse region from a URL or returns default.
 */
export function parseRegionFromUrl(url: string): Region {
  for (const region of VALID_REGIONS) {
    if (url.includes(`/api/chatsvc/${region}/`) ||
        url.includes(`/api/csa/${region}/`)) {
      return region;
    }
  }
  return 'amer'; // Default region
}

/** Substrate API endpoints. */
export const SUBSTRATE_API = {
  /** Full-text message search. */
  search: 'https://substrate.office.com/searchservice/api/v2/query',
  
  /** People and message typeahead suggestions. */
  suggestions: 'https://substrate.office.com/search/api/v1/suggestions',
  
  /** Frequent contacts list. */
  frequentContacts: 'https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache',
  
  /** People search. */
  peopleSearch: 'https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar',
} as const;

/** Chat service API endpoints. */
export const CHATSVC_API = {
  /** Get messages URL for a conversation. */
  messages: (region: Region, conversationId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`,
  
  /** Get conversation metadata URL. */
  conversation: (region: Region, conversationId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}`,
  
  /** Save/unsave message metadata URL. */
  messageMetadata: (region: Region, conversationId: string, messageId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/rcmetadata/${messageId}`,
} as const;

/** CSA (Chat Service Aggregator) API endpoints. */
export const CSA_API = {
  /** Conversation folders (favourites) URL. */
  conversationFolders: (region: Region) =>
    `https://teams.microsoft.com/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`,
} as const;

/** Common request headers for Teams API calls. */
export function getTeamsHeaders(): HeadersInit {
  return {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Origin': 'https://teams.microsoft.com',
    'Referer': 'https://teams.microsoft.com/',
  };
}

/** Headers for Bearer token authentication. */
export function getBearerHeaders(token: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authorization': `Bearer ${token}`,
  };
}

/** Headers for Skype token + Bearer authentication. */
export function getSkypeAuthHeaders(skypeToken: string, authToken: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${authToken}`,
  };
}

/** Headers for CSA API (Skype token + CSA Bearer). */
export function getCsaHeaders(skypeToken: string, csaToken: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${csaToken}`,
  };
}

/** Client version header for messaging API. */
export function getMessagingHeaders(skypeToken: string, authToken: string): HeadersInit {
  return {
    ...getSkypeAuthHeaders(skypeToken, authToken),
    'X-Ms-Client-Version': '1415/1.0.0.2025010401',
  };
}
