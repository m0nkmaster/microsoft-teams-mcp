/**
 * Authentication handling for Microsoft Teams.
 * Manages login detection and manual authentication flows.
 */

import type { Page, BrowserContext } from 'playwright';
import { saveSessionState } from './context.js';

const TEAMS_URL = 'https://teams.microsoft.com';

// URLs that indicate we're in a login flow
const LOGIN_URL_PATTERNS = [
  'login.microsoftonline.com',
  'login.live.com',
  'login.microsoft.com',
];

// Selectors that indicate successful authentication
const AUTH_SUCCESS_SELECTORS = [
  '[data-tid="app-bar"]',
  '[data-tid="search-box"]',
  'input[placeholder*="Search"]',
  '[data-tid="chat-list"]',
  '[data-tid="team-list"]',
];

export interface AuthStatus {
  isAuthenticated: boolean;
  isOnLoginPage: boolean;
  currentUrl: string;
}

/**
 * Checks if the current page URL indicates a login flow.
 */
function isLoginUrl(url: string): boolean {
  return LOGIN_URL_PATTERNS.some(pattern => url.includes(pattern));
}

/**
 * Checks if the page shows authenticated Teams content.
 */
async function hasAuthenticatedContent(page: Page): Promise<boolean> {
  for (const selector of AUTH_SUCCESS_SELECTORS) {
    try {
      const count = await page.locator(selector).count();
      if (count > 0) {
        return true;
      }
    } catch {
      // Selector not found, continue checking others
    }
  }
  return false;
}

/**
 * Gets the current authentication status.
 */
export async function getAuthStatus(page: Page): Promise<AuthStatus> {
  const currentUrl = page.url();
  const onLoginPage = isLoginUrl(currentUrl);
  
  // If on login page, definitely not authenticated
  if (onLoginPage) {
    return {
      isAuthenticated: false,
      isOnLoginPage: true,
      currentUrl,
    };
  }

  // If on Teams domain, check for authenticated content
  if (currentUrl.includes('teams.microsoft.com')) {
    const hasContent = await hasAuthenticatedContent(page);
    return {
      isAuthenticated: hasContent,
      isOnLoginPage: false,
      currentUrl,
    };
  }

  // Unknown state
  return {
    isAuthenticated: false,
    isOnLoginPage: false,
    currentUrl,
  };
}

/**
 * Navigates to Teams and checks authentication status.
 */
export async function navigateToTeams(page: Page): Promise<AuthStatus> {
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
  
  // Wait a moment for redirects to complete
  await page.waitForTimeout(2000);
  
  return getAuthStatus(page);
}

/**
 * Waits for the user to complete manual authentication.
 * Returns when authenticated or throws after timeout.
 * 
 * @param page - The page to monitor
 * @param context - Browser context for saving session
 * @param timeoutMs - Maximum time to wait (default: 5 minutes)
 * @param onProgress - Callback for progress updates
 */
export async function waitForManualLogin(
  page: Page,
  context: BrowserContext,
  timeoutMs: number = 5 * 60 * 1000,
  onProgress?: (message: string) => void
): Promise<void> {
  const startTime = Date.now();
  const log = onProgress ?? console.log;

  log('Waiting for manual login...');

  while (Date.now() - startTime < timeoutMs) {
    const status = await getAuthStatus(page);
    
    if (status.isAuthenticated) {
      log('Authentication successful!');
      
      // Wait for MSAL to refresh tokens in the background
      // This happens via JavaScript after the page loads
      log('Waiting for token refresh...');
      await page.waitForTimeout(5000);
      
      // Navigate to trigger any pending token operations
      await page.goto('https://teams.microsoft.com', { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(5000);
      
      // Save the session state with fresh tokens
      await saveSessionState(context);
      log('Session state saved.');
      
      return;
    }

    // Check every 2 seconds
    await page.waitForTimeout(2000);
  }

  throw new Error('Authentication timeout: user did not complete login within the allowed time');
}

/**
 * Performs a full authentication flow:
 * 1. Navigate to Teams
 * 2. Check if already authenticated
 * 3. If not, wait for manual login
 * 
 * @param page - The page to use
 * @param context - Browser context for session management
 * @param onProgress - Callback for progress updates
 */
export async function ensureAuthenticated(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Navigating to Teams...');
  const status = await navigateToTeams(page);

  if (status.isAuthenticated) {
    log('Already authenticated.');
    
    // Wait a moment for any token refresh to complete
    await page.waitForTimeout(3000);
    
    // Save the session state with potentially refreshed tokens
    await saveSessionState(context);
    
    return;
  }

  if (status.isOnLoginPage) {
    log('Login required. Please complete authentication in the browser window.');
    await waitForManualLogin(page, context, undefined, onProgress);
    
    // Navigate back to Teams after login (in case we're on a callback URL)
    await navigateToTeams(page);
  } else {
    // Unexpected state - might need manual intervention
    log('Unexpected page state. Waiting for authentication...');
    await waitForManualLogin(page, context, undefined, onProgress);
  }
}

/**
 * Forces a new login by clearing session and navigating to Teams.
 */
export async function forceNewLogin(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Starting fresh login...');
  
  // Clear cookies to force re-authentication
  await context.clearCookies();
  
  // Navigate and wait for login
  await navigateToTeams(page);
  await waitForManualLogin(page, context, undefined, onProgress);
}
