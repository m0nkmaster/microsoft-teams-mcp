/**
 * Secure session state storage.
 * 
 * Handles reading and writing session state with:
 * - Encryption at rest
 * - Restricted file permissions
 * - Automatic migration from plaintext
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { encrypt, decrypt, isEncrypted } from './crypto.js';
import { SESSION_EXPIRY_HOURS } from '../constants.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const PROJECT_ROOT = path.resolve(__dirname, '../..');
export const USER_DATA_DIR = path.join(PROJECT_ROOT, '.user-data');
export const SESSION_STATE_PATH = path.join(PROJECT_ROOT, 'session-state.json');
export const TOKEN_CACHE_PATH = path.join(PROJECT_ROOT, 'token-cache.json');

/** File permission mode: owner read/write only. */
const SECURE_FILE_MODE = 0o600;

/** Session state as stored by Playwright. */
export interface SessionState {
  cookies: Array<{
    name: string;
    value: string;
    domain?: string;
    path?: string;
    expires?: number;
    httpOnly?: boolean;
    secure?: boolean;
    sameSite?: 'Strict' | 'Lax' | 'None';
  }>;
  origins: Array<{
    origin: string;
    localStorage: Array<{ name: string; value: string }>;
  }>;
}

/** Token cache structure. */
export interface TokenCache {
  substrateToken: string;
  substrateTokenExpiry: number;
  extractedAt: number;
}

/**
 * Ensures the user data directory exists.
 */
export function ensureUserDataDir(): void {
  if (!fs.existsSync(USER_DATA_DIR)) {
    fs.mkdirSync(USER_DATA_DIR, { recursive: true, mode: 0o700 });
  }
}

/**
 * Writes data securely with encryption and file permissions.
 */
function writeSecure(filePath: string, data: unknown): void {
  const json = JSON.stringify(data, null, 2);
  const encrypted = encrypt(json);
  
  fs.writeFileSync(filePath, JSON.stringify(encrypted, null, 2), { 
    mode: SECURE_FILE_MODE,
    encoding: 'utf8',
  });
}

/**
 * Reads data securely, handling both encrypted and legacy plaintext.
 */
function readSecure<T>(filePath: string): T | null {
  if (!fs.existsSync(filePath)) {
    return null;
  }

  try {
    const content = fs.readFileSync(filePath, 'utf8');
    const parsed = JSON.parse(content);

    // Check if this is encrypted data
    if (isEncrypted(parsed)) {
      const decrypted = decrypt(parsed);
      return JSON.parse(decrypted) as T;
    }

    // Legacy plaintext - migrate to encrypted
    writeSecure(filePath, parsed);
    return parsed as T;

  } catch (error) {
    // If decryption fails (different machine, corrupted), return null
    console.error(`Failed to read ${filePath}:`, error instanceof Error ? error.message : error);
    return null;
  }
}

/**
 * Checks if session state file exists.
 */
export function hasSessionState(): boolean {
  return fs.existsSync(SESSION_STATE_PATH);
}

/**
 * Reads the session state.
 */
export function readSessionState(): SessionState | null {
  return readSecure<SessionState>(SESSION_STATE_PATH);
}

/**
 * Writes the session state securely.
 */
export function writeSessionState(state: SessionState): void {
  writeSecure(SESSION_STATE_PATH, state);
}

/**
 * Deletes the session state file.
 */
export function clearSessionState(): void {
  if (fs.existsSync(SESSION_STATE_PATH)) {
    fs.unlinkSync(SESSION_STATE_PATH);
  }
}

/**
 * Gets the age of the session state in hours.
 */
export function getSessionAge(): number | null {
  if (!hasSessionState()) {
    return null;
  }

  const stats = fs.statSync(SESSION_STATE_PATH);
  const ageMs = Date.now() - stats.mtimeMs;
  return ageMs / (1000 * 60 * 60);
}

/**
 * Checks if session is likely expired (>12 hours old).
 */
export function isSessionLikelyExpired(): boolean {
  const age = getSessionAge();
  if (age === null) return true;
  return age > SESSION_EXPIRY_HOURS;
}

/**
 * Reads the token cache.
 */
export function readTokenCache(): TokenCache | null {
  return readSecure<TokenCache>(TOKEN_CACHE_PATH);
}

/**
 * Writes the token cache securely.
 */
export function writeTokenCache(cache: TokenCache): void {
  writeSecure(TOKEN_CACHE_PATH, cache);
}

/**
 * Clears the token cache.
 */
export function clearTokenCache(): void {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
}

/**
 * Gets the Teams origin from session state.
 */
export function getTeamsOrigin(state: SessionState): SessionState['origins'][number] | null {
  return state.origins?.find(o => o.origin === 'https://teams.microsoft.com') ?? null;
}
