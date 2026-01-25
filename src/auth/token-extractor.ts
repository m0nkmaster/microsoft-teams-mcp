/**
 * Token extraction from session state.
 * 
 * Extracts various authentication tokens from Playwright's saved session state.
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

/**
 * Decodes a JWT token's payload without verifying the signature.
 * Returns null if the token is invalid.
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
 * Gets the expiry date from a JWT token.
 */
function getJwtExpiry(token: string): Date | null {
  const payload = decodeJwtPayload(token);
  if (!payload?.exp || typeof payload.exp !== 'number') return null;
  return new Date(payload.exp * 1000);
}

/** Substrate search token information. */
export interface SubstrateTokenInfo {
  token: string;
  expiry: Date;
}

/** Teams API token information. */
export interface TeamsTokenInfo {
  token: string;
  expiry: Date;
  userMri: string;
}

/** Message authentication information (cookies). */
export interface MessageAuthInfo {
  skypeToken: string;
  authToken: string;
  userMri: string;
}

/**
 * Extracts the Substrate search token from session state.
 */
export function extractSubstrateToken(state?: SessionState): SubstrateTokenInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
  if (!teamsOrigin) return null;

  for (const item of teamsOrigin.localStorage) {
    try {
      const val = JSON.parse(item.value);
      if (val.target?.includes('substrate.office.com/search/SubstrateSearch')) {
        const token = val.secret;
        if (!token || typeof token !== 'string') continue;

        const expiry = getJwtExpiry(token);
        if (expiry) {
          return { token, expiry };
        }
      }
    } catch {
      continue;
    }
  }

  return null;
}

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

/**
 * Extracts the Teams chat API token from session state.
 */
export function extractTeamsToken(state?: SessionState): TeamsTokenInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
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

      const payload = decodeJwtPayload(secret);
      if (!payload?.exp || typeof payload.exp !== 'number') continue;
      const tokenExpiry = new Date(payload.exp * 1000);

      // Extract user MRI from any token
      if (typeof payload.oid === 'string' && !userMri) {
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
    const substrateInfo = extractSubstrateToken(sessionState);
    if (substrateInfo) {
      const payload = decodeJwtPayload(substrateInfo.token);
      if (typeof payload?.oid === 'string') {
        userMri = `8:orgid:${payload.oid}`;
      }
    }
  }

  // Prefer chatsvc token, fallback to skype token
  const token = chatToken || skypeToken;
  const expiry = chatToken ? chatTokenExpiry : skypeTokenExpiry;

  if (token && expiry && userMri && expiry.getTime() > Date.now()) {
    return { token, expiry, userMri };
  }

  return null;
}

/**
 * Extracts authentication info needed for messaging API (uses cookies).
 */
export function extractMessageAuth(state?: SessionState): MessageAuthInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  let skypeToken: string | null = null;
  let authToken: string | null = null;
  let userMri: string | null = null;

  // Extract tokens from cookies
  for (const cookie of sessionState.cookies || []) {
    if (cookie.name === 'skypetoken_asm' && cookie.domain?.includes('teams.microsoft.com')) {
      skypeToken = cookie.value;
    }
    if (cookie.name === 'authtoken' && cookie.domain?.includes('teams.microsoft.com')) {
      authToken = decodeURIComponent(cookie.value);
      if (authToken.startsWith('Bearer=')) {
        authToken = authToken.substring(7);
      }
    }
  }

  // Get userMri from skypeToken payload
  if (skypeToken) {
    const payload = decodeJwtPayload(skypeToken);
    if (typeof payload?.skypeid === 'string') {
      userMri = payload.skypeid;
    }
  }

  // Fallback to extracting userMri from authToken
  if (!userMri && authToken) {
    const payload = decodeJwtPayload(authToken);
    if (typeof payload?.oid === 'string') {
      userMri = `8:orgid:${payload.oid}`;
    }
  }

  if (skypeToken && authToken && userMri) {
    return { skypeToken, authToken, userMri };
  }

  return null;
}

/**
 * Extracts the CSA token for the conversationFolders API.
 */
export function extractCsaToken(state?: SessionState): string | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  for (const origin of sessionState.origins || []) {
    for (const item of origin.localStorage || []) {
      if (item.name.includes('chatsvcagg.teams.microsoft.com') && !item.name.startsWith('tmp.')) {
        try {
          const data = JSON.parse(item.value) as { secret?: string };
          if (data.secret) {
            return data.secret;
          }
        } catch {
          // Ignore parse errors
        }
      }
    }
  }

  return null;
}

/**
 * Gets the current user's profile from cached JWT tokens.
 */
export function getUserProfile(state?: SessionState): UserProfile | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
  if (!teamsOrigin) return null;

  for (const item of teamsOrigin.localStorage) {
    try {
      const val = JSON.parse(item.value);
      if (!val.secret || typeof val.secret !== 'string') continue;
      if (!val.secret.startsWith('ey')) continue;

      const payload = decodeJwtPayload(val.secret);
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
 */
export function getUserDisplayName(state?: SessionState): string | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
  if (!teamsOrigin) return null;

  for (const item of teamsOrigin.localStorage) {
    try {
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
  const teamsToken = extractTeamsToken(sessionState);
  if (teamsToken) {
    const payload = decodeJwtPayload(teamsToken.token);
    if (typeof payload?.name === 'string') return payload.name;
  }

  return null;
}

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
