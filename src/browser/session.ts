/**
 * Session persistence utilities.
 * Handles saving and restoring browser session state.
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const PROJECT_ROOT = path.resolve(__dirname, '../..');
export const USER_DATA_DIR = path.join(PROJECT_ROOT, '.user-data');
export const SESSION_STATE_PATH = path.join(PROJECT_ROOT, 'session-state.json');

/**
 * Ensures the user data directory exists.
 */
export function ensureUserDataDir(): void {
  if (!fs.existsSync(USER_DATA_DIR)) {
    fs.mkdirSync(USER_DATA_DIR, { recursive: true });
  }
}

/**
 * Checks if a saved session state exists.
 */
export function hasSessionState(): boolean {
  return fs.existsSync(SESSION_STATE_PATH);
}

/**
 * Deletes the saved session state.
 */
export function clearSessionState(): void {
  if (fs.existsSync(SESSION_STATE_PATH)) {
    fs.unlinkSync(SESSION_STATE_PATH);
  }
}

/**
 * Gets the age of the session state file in hours.
 * Returns null if no session state exists.
 */
export function getSessionAge(): number | null {
  if (!hasSessionState()) {
    return null;
  }

  const stats = fs.statSync(SESSION_STATE_PATH);
  const ageMs = Date.now() - stats.mtimeMs;
  return ageMs / (1000 * 60 * 60); // Convert to hours
}

/**
 * Checks if the session is likely expired based on file age.
 * Sessions older than 12 hours are considered potentially expired.
 */
export function isSessionLikelyExpired(): boolean {
  const age = getSessionAge();
  if (age === null) {
    return true;
  }
  return age > 12; // 12 hours
}

/**
 * Checks if the tokens in the session state are expired.
 * Returns true if no tokens found or if the Substrate search token is expired.
 */
export function areTokensExpired(): boolean {
  if (!hasSessionState()) {
    return true;
  }

  try {
    const state = JSON.parse(fs.readFileSync(SESSION_STATE_PATH, 'utf8'));
    const teamsOrigin = state.origins?.find((o: { origin: string }) => 
      o.origin === 'https://teams.microsoft.com'
    );

    if (!teamsOrigin) return true;

    // Look for Substrate search token
    for (const item of teamsOrigin.localStorage) {
      try {
        const val = JSON.parse(item.value);
        if (val.target?.includes('substrate.office.com/search/SubstrateSearch')) {
          const parts = val.secret.split('.');
          if (parts.length === 3) {
            const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
            const expiry = new Date(payload.exp * 1000);
            return expiry.getTime() <= Date.now();
          }
        }
      } catch {
        continue;
      }
    }
  } catch {
    return true;
  }

  return true; // No token found = expired
}
