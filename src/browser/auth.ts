/**
 * Authentication handling for Microsoft Teams.
 * Manages login detection and manual authentication flows.
 */

import type { Page, BrowserContext } from 'playwright';
import { saveSessionState } from './context.js';
import {
  OVERLAY_STEP_PAUSE_MS,
  OVERLAY_COMPLETE_PAUSE_MS,
} from '../constants.js';

const TEAMS_URL = 'https://teams.microsoft.com';

// ─────────────────────────────────────────────────────────────────────────────
// Progress Overlay UI
// ─────────────────────────────────────────────────────────────────────────────

const PROGRESS_OVERLAY_ID = 'mcp-login-progress-overlay';

/** Phases for the login progress overlay. */
type OverlayPhase = 'signed-in' | 'acquiring' | 'saving' | 'complete' | 'refreshing' | 'error';

/** Content for each overlay phase. */
const OVERLAY_CONTENT: Record<OverlayPhase, { message: string; detail: string }> = {
  'signed-in': {
    message: "You're signed in!",
    detail: 'Setting up your connection to Teams...',
  },
  'acquiring': {
    message: 'Acquiring permissions...',
    detail: 'Getting access to search and messages...',
  },
  'saving': {
    message: 'Saving your session...',
    detail: "So you won't need to log in again.",
  },
  'complete': {
    message: 'All done!',
    detail: 'This window will close automatically.',
  },
  'refreshing': {
    message: 'Refreshing your session...',
    detail: 'Updating your access tokens...',
  },
  'error': {
    message: 'Something went wrong',
    detail: 'Please try again or check the console for details.',
  },
};

/**
 * Shows a progress overlay for a specific phase.
 * Handles injection, content, and optional pause.
 * Failures are silently ignored - the overlay is purely cosmetic.
 */
async function showLoginProgress(
  page: Page,
  phase: OverlayPhase,
  options: { pause?: boolean } = {}
): Promise<void> {
  const content = OVERLAY_CONTENT[phase];
  const isComplete = phase === 'complete';
  const isError = phase === 'error';

  try {
    await page.evaluate(({ id, message, detail, complete, error }) => {
      // Remove existing overlay if present
      const existing = document.getElementById(id);
      if (existing) {
        existing.remove();
      }

      // Create overlay container
      const overlay = document.createElement('div');
      overlay.id = id;
      Object.assign(overlay.style, {
        position: 'fixed',
        top: '0',
        left: '0',
        right: '0',
        bottom: '0',
        background: 'rgba(0, 0, 0, 0.7)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: '999999',
        fontFamily: "'Segoe UI', system-ui, sans-serif",
      });

      // Create modal card
      const modal = document.createElement('div');
      Object.assign(modal.style, {
        background: 'white',
        borderRadius: '12px',
        padding: '40px 48px',
        maxWidth: '420px',
        textAlign: 'center',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.3)',
      });

      // Create icon
      const icon = document.createElement('div');
      const iconBg = error ? '#c42b1c' : complete ? '#107c10' : '#5b5fc7';
      const iconText = error ? '✕' : complete ? '✓' : '⋯';
      Object.assign(icon.style, {
        width: '64px',
        height: '64px',
        margin: '0 auto 24px',
        borderRadius: '50%',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: '32px',
        background: iconBg,
        color: 'white',
      });
      icon.textContent = iconText;

      // Create title
      const title = document.createElement('h2');
      Object.assign(title.style, {
        margin: '0 0 12px',
        fontSize: '20px',
        fontWeight: '600',
        color: '#242424',
      });
      title.textContent = message;

      // Create detail text
      const detailEl = document.createElement('p');
      Object.assign(detailEl.style, {
        margin: '0',
        fontSize: '14px',
        color: '#616161',
        lineHeight: '1.5',
      });
      detailEl.textContent = detail;

      // Assemble and append
      modal.appendChild(icon);
      modal.appendChild(title);
      modal.appendChild(detailEl);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
    }, {
      id: PROGRESS_OVERLAY_ID,
      message: content.message,
      detail: content.detail,
      complete: isComplete,
      error: isError,
    });

    // Pause if requested (for steps that need user to see the message)
    if (options.pause) {
      const pauseMs = isComplete ? OVERLAY_COMPLETE_PAUSE_MS : OVERLAY_STEP_PAUSE_MS;
      await page.waitForTimeout(pauseMs);
    }
  } catch {
    // Overlay is cosmetic - don't fail login if it can't be shown
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Authentication Detection
// ─────────────────────────────────────────────────────────────────────────────

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

// Search input selectors - must match those in teams/search.ts
const SEARCH_INPUT_SELECTORS = [
  '[data-tid="searchInputField"]',
  '[data-tid="app-search-input"]',
  'input[data-tid*="search"]',
  'input[placeholder*="Search"]',
];

const SEARCH_BUTTON_SELECTORS = [
  '[data-tid="search-box"]',
  '[data-tid="search-button"]',
  '[data-tid="app-bar-search"]',
  '[role="search"] input',
];

/**
 * Triggers MSAL to acquire the Substrate token.
 * 
 * MSAL only acquires tokens for specific scopes when the app makes API calls
 * requiring those scopes. The Substrate API is only used for search, so we
 * perform a minimal search ("is:Messages") to trigger token acquisition.
 */
async function triggerTokenAcquisition(
  page: Page,
  log: (msg: string) => void
): Promise<void> {
  log('Triggering token acquisition...');

  try {
    // Wait for the app to be ready
    await page.waitForTimeout(3000);

    // Try to find a search input directly
    let searchInput = null;
    for (const selector of SEARCH_INPUT_SELECTORS) {
      const loc = page.locator(selector).first();
      if (await loc.isVisible().catch(() => false)) {
        searchInput = loc;
        break;
      }
    }

    // If no input visible, try clicking a search button first
    if (!searchInput) {
      for (const selector of SEARCH_BUTTON_SELECTORS) {
        const btn = page.locator(selector).first();
        if (await btn.isVisible().catch(() => false)) {
          await btn.click();
          await page.waitForTimeout(1000);
          
          // Now look for the input again
          for (const inputSelector of SEARCH_INPUT_SELECTORS) {
            const loc = page.locator(inputSelector).first();
            if (await loc.isVisible().catch(() => false)) {
              searchInput = loc;
              break;
            }
          }
          break;
        }
      }
    }

    if (searchInput) {
      // Use a filter syntax that looks like a system command
      await searchInput.fill('is:Messages');
      await page.keyboard.press('Enter');

      // Wait for the search API call to complete
      log('Waiting for search API...');
      await page.waitForTimeout(5000);

      // Close search
      await page.keyboard.press('Escape');
      await page.waitForTimeout(1000);

      log('Token acquisition complete.');
    } else {
      log('Search UI not found, waiting for background refresh...');
      await page.waitForTimeout(5000);
    }
  } catch (error) {
    log(`Token acquisition warning: ${error instanceof Error ? error.message : String(error)}`);
    await page.waitForTimeout(3000);
  }
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

      // Show progress through login steps
      await showLoginProgress(page, 'signed-in', { pause: true });
      await showLoginProgress(page, 'acquiring');

      // Trigger a search to cause MSAL to acquire the Substrate token
      await triggerTokenAcquisition(page, log);

      await showLoginProgress(page, 'saving');

      // Save the session state with fresh tokens
      await saveSessionState(context);
      log('Session state saved.');

      await showLoginProgress(page, 'complete', { pause: true });

      return;
    }

    // Check every 2 seconds
    await page.waitForTimeout(2000);
  }

  // Show error overlay before throwing
  await showLoginProgress(page, 'error', { pause: true });

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
 * @param showOverlay - Whether to show progress overlay (default: true for visible browsers)
 */
export async function ensureAuthenticated(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
  showOverlay: boolean = true
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Navigating to Teams...');
  const status = await navigateToTeams(page);

  if (status.isAuthenticated) {
    log('Already authenticated.');

    if (showOverlay) {
      await showLoginProgress(page, 'refreshing');
    }

    // Trigger a search to cause MSAL to acquire/refresh the Substrate token
    await triggerTokenAcquisition(page, log);

    if (showOverlay) {
      await showLoginProgress(page, 'saving');
    }

    // Save the session state with fresh tokens
    await saveSessionState(context);

    if (showOverlay) {
      await showLoginProgress(page, 'complete', { pause: true });
    }

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
