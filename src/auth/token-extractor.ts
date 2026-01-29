/**
 * Token extraction from session state.
 * 
 * Extracts various authentication tokens from Playwright's saved session state.
 * Teams stores MSAL tokens in localStorage; we parse these to get bearer tokens
 * for various APIs (Substrate search, chatsvc messaging, etc.).
 */

import {
  readSessionState,
  readTokenCache,
  writeTokenCache,
  clearTokenCache,
  getTeamsOrigin,
  type SessionState,
  type TokenCache,
} from './session-store.js';
import { parseJwtProfile, type UserProfile } from '../utils/parsers.js';
import { MRI_ORGID_PREFIX } from '../constants.js';

// ============================================================================
// JWT Utilities
// ============================================================================

/**
 * Decodes a JWT token's payload without verifying the signature.
 */
function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    return JSON.parse(Buffer.from(parts[1], 'base64').toString());
  } catch {
    return null;
  }
}

/**
 * Gets the expiry date from a JWT token's `exp` claim.
 */
function getJwtExpiry(token: string): Date | null {
  const payload = decodeJwtPayload(token);
  if (!payload?.exp || typeof payload.exp !== 'number') return null;
  return new Date(payload.exp * 1000);
}

/**
 * Checks if a string looks like a JWT (starts with 'ey').
 */
function isJwtToken(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith('ey');
}

// ============================================================================
// Session Helpers
// ============================================================================

/**
 * Resolves session state and Teams origin in one call.
 * Many functions need both, so this reduces boilerplate.
 */
function getTeamsLocalStorage(state?: SessionState): Array<{ name: string; value: string }> | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
  return teamsOrigin?.localStorage ?? null;
}

// ============================================================================
// Types
// ============================================================================

/** Substrate search token (for search/people APIs). */
export interface SubstrateTokenInfo {
  token: string;
  expiry: Date;
}

/** Teams chat API token (for chatsvc). */
export interface TeamsTokenInfo {
  token: string;
  expiry: Date;
  userMri: string;
}

/** Cookie-based auth for messaging APIs. */
export interface MessageAuthInfo {
  skypeToken: string;
  authToken: string;
  userMri: string;
}

// ============================================================================
// Token Extraction
// ============================================================================

/**
 * Extracts the Substrate search token from session state.
 * This token is used for search and people APIs.
 */
export function extractSubstrateToken(state?: SessionState): SubstrateTokenInfo | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  // Collect all valid Substrate tokens and pick the one with longest expiry
  let bestToken: SubstrateTokenInfo | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      
      // Look for Substrate search tokens by target scope
      // Match both old format (substrate.office.com/search/SubstrateSearch)
      // and new format (substrate.office.com/SubstrateSearch-Internal.ReadWrite)
      const target = entry.target as string | undefined;
      if (!target?.includes('substrate.office.com')) continue;
      if (!target.includes('SubstrateSearch')) continue;

      if (!isJwtToken(entry.secret)) continue;

      const expiry = getJwtExpiry(entry.secret);
      if (!expiry) continue;

      // Skip expired tokens
      if (expiry.getTime() <= Date.now()) continue;

      // Keep the token with longest remaining validity
      if (!bestToken || expiry.getTime() > bestToken.expiry.getTime()) {
        bestToken = { token: entry.secret, expiry };
      }
    } catch {
      continue;
    }
  }

  return bestToken;
}

// ============================================================================
// Cached Token Access
// ============================================================================

/**
 * Gets a valid Substrate token, either from cache or by extracting from session.
 */
export function getValidSubstrateToken(): string | null {
  // Try cache first
  const cache = readTokenCache();
  if (cache && cache.substrateTokenExpiry > Date.now()) {
    return cache.substrateToken;
  }

  // Extract from session
  const extracted = extractSubstrateToken();
  if (!extracted) return null;

  // Check if not expired
  if (extracted.expiry.getTime() <= Date.now()) {
    return null;
  }

  // Cache the token
  const newCache: TokenCache = {
    substrateToken: extracted.token,
    substrateTokenExpiry: extracted.expiry.getTime(),
    extractedAt: Date.now(),
  };
  writeTokenCache(newCache);

  return extracted.token;
}

/**
 * Checks if we have a valid Substrate token.
 */
export function hasValidSubstrateToken(): boolean {
  return getValidSubstrateToken() !== null;
}

/**
 * Gets Substrate token status for diagnostics.
 */
export function getSubstrateTokenStatus(): {
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

/** Candidate token found during extraction. */
interface TokenCandidate {
  token: string;
  expiry: Date;
  userMri?: string;
}

/**
 * Extracts the Teams chat API token from session state.
 * 
 * Teams stores multiple tokens for different services. We prefer:
 * 1. chatsvcagg.teams.microsoft.com (primary chat API)
 * 2. api.spaces.skype.com (fallback)
 */
export function extractTeamsToken(state?: SessionState): TeamsTokenInfo | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  let chatsvcCandidate: TokenCandidate | null = null;
  let skypeCandidate: TokenCandidate | null = null;
  let userMri: string | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (!entry.target || !isJwtToken(entry.secret)) continue;

      const payload = decodeJwtPayload(entry.secret);
      if (!payload?.exp || typeof payload.exp !== 'number') continue;

      const expiry = new Date(payload.exp * 1000);

      // Capture user MRI from any token's oid claim
      if (typeof payload.oid === 'string' && !userMri) {
        userMri = `${MRI_ORGID_PREFIX}${payload.oid}`;
      }

      // Track best candidate for each service
      if (entry.target.includes('chatsvcagg.teams.microsoft.com')) {
        if (!chatsvcCandidate || expiry > chatsvcCandidate.expiry) {
          chatsvcCandidate = { token: entry.secret, expiry };
        }
      } else if (entry.target.includes('api.spaces.skype.com')) {
        if (!skypeCandidate || expiry > skypeCandidate.expiry) {
          skypeCandidate = { token: entry.secret, expiry };
        }
      }
    } catch {
      continue;
    }
  }

  // Fallback: extract userMri from Substrate token if not found
  if (!userMri) {
    userMri = extractUserMriFromSubstrate(state);
  }

  // Prefer chatsvc, fall back to skype
  const best = chatsvcCandidate ?? skypeCandidate;
  if (!best || !userMri || best.expiry.getTime() <= Date.now()) {
    return null;
  }

  return { token: best.token, expiry: best.expiry, userMri };
}

/**
 * Extracts user MRI from the Substrate token's oid claim.
 */
function extractUserMriFromSubstrate(state?: SessionState): string | null {
  const substrateInfo = extractSubstrateToken(state);
  if (!substrateInfo) return null;

  const payload = decodeJwtPayload(substrateInfo.token);
  if (typeof payload?.oid === 'string') {
    return `${MRI_ORGID_PREFIX}${payload.oid}`;
  }
  return null;
}

/**
 * Extracts authentication info needed for messaging API.
 * Unlike other APIs, messaging uses cookies rather than localStorage tokens.
 */
export function extractMessageAuth(state?: SessionState): MessageAuthInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const cookies = sessionState.cookies ?? [];
  const teamsCookies = cookies.filter(c => c.domain?.includes('teams.microsoft.com'));

  // Extract the two required cookies
  const skypeToken = teamsCookies.find(c => c.name === 'skypetoken_asm')?.value ?? null;
  const rawAuthToken = teamsCookies.find(c => c.name === 'authtoken')?.value ?? null;
  
  if (!skypeToken || !rawAuthToken) return null;

  // Decode authtoken (URL-encoded, may have 'Bearer=' prefix)
  let authToken = decodeURIComponent(rawAuthToken);
  if (authToken.startsWith('Bearer=')) {
    authToken = authToken.substring(7);
  }

  // Extract userMri from skypeToken's skypeid claim, or fall back to authToken's oid
  const userMri = extractMriFromSkypeToken(skypeToken) 
    ?? extractMriFromAuthToken(authToken);

  if (!userMri) return null;

  return { skypeToken, authToken, userMri };
}

/**
 * Gets messaging token status for diagnostics.
 * The skypetoken_asm cookie is a JWT with an exp claim.
 */
export function getMessageAuthStatus(): {
  hasToken: boolean;
  expiresAt?: string;
  minutesRemaining?: number;
} {
  const sessionState = readSessionState();
  if (!sessionState) {
    return { hasToken: false };
  }

  const cookies = sessionState.cookies ?? [];
  const skypeToken = cookies.find(
    c => c.domain?.includes('teams.microsoft.com') && c.name === 'skypetoken_asm'
  )?.value;

  if (!skypeToken) {
    return { hasToken: false };
  }

  const expiry = getJwtExpiry(skypeToken);
  if (!expiry) {
    // Token exists but can't parse expiry - assume valid
    return { hasToken: true };
  }

  const now = Date.now();
  const expiryMs = expiry.getTime();

  return {
    hasToken: expiryMs > now,
    expiresAt: expiry.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - now) / 1000 / 60)),
  };
}

function extractMriFromSkypeToken(token: string): string | null {
  const payload = decodeJwtPayload(token);
  return typeof payload?.skypeid === 'string' ? payload.skypeid : null;
}

function extractMriFromAuthToken(token: string): string | null {
  const payload = decodeJwtPayload(token);
  return typeof payload?.oid === 'string' ? `${MRI_ORGID_PREFIX}${payload.oid}` : null;
}

/**
 * Extracts the CSA token for the conversationFolders API.
 * This searches all origins, not just teams.microsoft.com.
 */
export function extractCsaToken(state?: SessionState): string | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  for (const origin of sessionState.origins ?? []) {
    for (const item of origin.localStorage ?? []) {
      // Skip temporary entries, look for chatsvcagg tokens
      if (item.name.startsWith('tmp.')) continue;
      if (!item.name.includes('chatsvcagg.teams.microsoft.com')) continue;

      try {
        const entry = JSON.parse(item.value) as { secret?: string };
        if (entry.secret) return entry.secret;
      } catch {
        // Ignore parse errors
      }
    }
  }

  return null;
}

// ============================================================================
// User Profile
// ============================================================================

/**
 * Gets the current user's profile from cached JWT tokens.
 */
export function getUserProfile(state?: SessionState): UserProfile | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (!isJwtToken(entry.secret)) continue;

      const payload = decodeJwtPayload(entry.secret);
      if (payload) {
        const profile = parseJwtProfile(payload);
        if (profile) return profile;
      }
    } catch {
      continue;
    }
  }

  return null;
}

/**
 * Gets user's display name from session state.
 * Searches localStorage entries first, then falls back to JWT claims.
 */
export function getUserDisplayName(state?: SessionState): string | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  // First pass: look for explicit displayName in localStorage
  for (const item of localStorage) {
    // Quick filter before parsing
    if (!item.value?.includes('displayName') && !item.value?.includes('givenName')) {
      continue;
    }

    try {
      const entry = JSON.parse(item.value);
      if (entry.displayName) return entry.displayName;
      if (entry.name?.displayName) return entry.name.displayName;
    } catch {
      continue;
    }
  }

  // Fallback: extract from Teams token's name claim
  const teamsToken = extractTeamsToken(state);
  if (teamsToken) {
    const payload = decodeJwtPayload(teamsToken.token);
    if (typeof payload?.name === 'string') return payload.name;
  }

  return null;
}

// ============================================================================
// Token Status Checks
// ============================================================================

/**
 * Checks if tokens in session state are expired.
 */
export function areTokensExpired(state?: SessionState): boolean {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return true;

  const substrate = extractSubstrateToken(sessionState);
  return !substrate || substrate.expiry.getTime() <= Date.now();
}

// Re-export clearTokenCache for convenience
export { clearTokenCache };
