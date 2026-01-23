/**
 * Playwright browser context management.
 * Creates and manages browser contexts with session persistence.
 */

import { chromium, type Browser, type BrowserContext, type Page } from 'playwright';
import { 
  ensureUserDataDir, 
  hasSessionState, 
  SESSION_STATE_PATH,
  isSessionLikelyExpired,
  areTokensExpired,
  clearSessionState
} from './session.js';

export interface BrowserManager {
  browser: Browser;
  context: BrowserContext;
  page: Page;
  isNewSession: boolean;
}

export interface CreateBrowserOptions {
  headless?: boolean;
  viewport?: { width: number; height: number };
}

const DEFAULT_OPTIONS: Required<CreateBrowserOptions> = {
  headless: true,
  viewport: { width: 1280, height: 800 },
};

/**
 * Creates a browser context with optional session state restoration.
 * 
 * @param options - Browser configuration options
 * @returns Browser manager with browser, context, and page
 */
export async function createBrowserContext(
  options: CreateBrowserOptions = {}
): Promise<BrowserManager> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  
  ensureUserDataDir();

  const browser = await chromium.launch({
    headless: opts.headless,
  });

  const hasSession = hasSessionState();
  const sessionExpired = isSessionLikelyExpired();
  const tokensExpired = areTokensExpired();
  
  // Restore session if we have one and it's not ancient
  // Session cookies help with re-auth even if MSAL tokens are expired
  const shouldRestoreSession = hasSession && !sessionExpired;

  let context: BrowserContext;
  
  if (shouldRestoreSession) {
    try {
      context = await browser.newContext({
        storageState: SESSION_STATE_PATH,
        viewport: opts.viewport,
      });
    } catch (error) {
      console.warn('Failed to restore session state, starting fresh:', error);
      context = await browser.newContext({
        viewport: opts.viewport,
      });
    }
  } else {
    context = await browser.newContext({
      viewport: opts.viewport,
    });
  }

  const page = await context.newPage();

  return {
    browser,
    context,
    page,
    isNewSession: !shouldRestoreSession,
  };
}

/**
 * Saves the current browser context's session state.
 */
export async function saveSessionState(context: BrowserContext): Promise<void> {
  await context.storageState({ path: SESSION_STATE_PATH });
}

/**
 * Closes the browser and optionally saves session state.
 */
export async function closeBrowser(
  manager: BrowserManager, 
  saveSession: boolean = true
): Promise<void> {
  if (saveSession) {
    await saveSessionState(manager.context);
  }
  await manager.context.close();
  await manager.browser.close();
}
