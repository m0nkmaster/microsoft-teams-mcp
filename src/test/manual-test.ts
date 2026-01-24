#!/usr/bin/env npx tsx
/**
 * Manual testing script for Teams MCP functionality.
 * Runs through the core features to verify they work.
 * 
 * Usage:
 *   npm run test:manual
 *   npm run test:manual -- --search "your query"
 *   npm run test:manual -- --headless
 */

import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { getAuthStatus, ensureAuthenticated } from '../browser/auth.js';
import { searchTeams } from '../teams/search.js';
import { hasSessionState, getSessionAge } from '../auth/session-store.js';

interface TestOptions {
  headless: boolean;
  searchQuery?: string;
}

function parseArgs(): TestOptions {
  const args = process.argv.slice(2);
  const options: TestOptions = {
    headless: false,
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--headless') {
      options.headless = true;
    } else if (args[i] === '--search' && args[i + 1]) {
      options.searchQuery = args[i + 1];
      i++;
    }
  }

  return options;
}

function log(message: string, indent = 0): void {
  const prefix = '  '.repeat(indent);
  console.log(`${prefix}${message}`);
}

function logSection(title: string): void {
  console.log('\n' + '─'.repeat(50));
  console.log(`  ${title}`);
  console.log('─'.repeat(50));
}

async function testSessionState(): Promise<boolean> {
  logSection('Session State');
  
  const hasSession = hasSessionState();
  const sessionAge = getSessionAge();
  
  log(`Session exists: ${hasSession ? '✅ Yes' : '❌ No'}`);
  
  if (sessionAge !== null) {
    const ageHours = sessionAge.toFixed(1);
    const isOld = sessionAge > 12;
    log(`Session age: ${ageHours} hours ${isOld ? '⚠️ (may be expired)' : '✅'}`);
  }
  
  return hasSession;
}

async function testBrowserContext(headless: boolean): Promise<Awaited<ReturnType<typeof createBrowserContext>> | null> {
  logSection('Browser Context');
  
  log(`Creating browser (headless: ${headless})...`);
  
  try {
    const manager = await createBrowserContext({ headless });
    log(`Browser launched: ✅`);
    log(`New session: ${manager.isNewSession ? 'Yes (will need login)' : 'No (restored from saved)'}`);
    return manager;
  } catch (error) {
    log(`Browser launch failed: ❌`);
    log(`Error: ${error instanceof Error ? error.message : String(error)}`, 1);
    return null;
  }
}

async function testAuthentication(
  manager: Awaited<ReturnType<typeof createBrowserContext>>
): Promise<boolean> {
  logSection('Authentication');
  
  log('Checking authentication status...');
  
  try {
    await ensureAuthenticated(
      manager.page,
      manager.context,
      (msg) => log(`  ${msg}`)
    );
    
    const status = await getAuthStatus(manager.page);
    log(`Authenticated: ${status.isAuthenticated ? '✅ Yes' : '❌ No'}`);
    log(`Current URL: ${status.currentUrl}`, 1);
    
    return status.isAuthenticated;
  } catch (error) {
    log(`Authentication failed: ❌`);
    log(`Error: ${error instanceof Error ? error.message : String(error)}`, 1);
    return false;
  }
}

async function testSearch(
  manager: Awaited<ReturnType<typeof createBrowserContext>>,
  query: string
): Promise<boolean> {
  logSection('Search');
  
  log(`Searching for: "${query}"...`);
  
  try {
    const results = await searchTeams(manager.page, query, {
      maxResults: 10,
      waitMs: 8000,
    });
    
    log(`Results found: ${results.length}`);
    
    if (results.length > 0) {
      log('Sample results:', 1);
      for (const result of results.slice(0, 3)) {
        const preview = result.content.substring(0, 80).replace(/\n/g, ' ');
        log(`• ${preview}${result.content.length > 80 ? '...' : ''}`, 2);
        if (result.sender) {
          log(`  From: ${result.sender}`, 2);
        }
      }
    }
    
    return results.length > 0;
  } catch (error) {
    log(`Search failed: ❌`);
    log(`Error: ${error instanceof Error ? error.message : String(error)}`, 1);
    return false;
  }
}

async function runTests(): Promise<void> {
  console.log('\n🧪 Teams MCP Manual Test');
  console.log('========================\n');
  
  const options = parseArgs();
  
  if (options.headless) {
    log('Running in headless mode');
  } else {
    log('Running with visible browser (use --headless to run headless)');
  }
  
  // Test 1: Session state
  const hasSession = await testSessionState();
  
  if (!hasSession && options.headless) {
    log('\n⚠️  No session found. Cannot run headless without a saved session.');
    log('   Run without --headless first to log in, or run: npm run research');
    process.exit(1);
  }
  
  // Test 2: Browser context
  const manager = await testBrowserContext(options.headless);
  if (!manager) {
    process.exit(1);
  }
  
  try {
    // Test 3: Authentication
    const isAuthenticated = await testAuthentication(manager);
    
    if (!isAuthenticated) {
      log('\n⚠️  Not authenticated. Please log in manually in the browser window.');
      log('   Waiting for authentication...');
      // The ensureAuthenticated call above should have handled this
    }
    
    // Test 4: Search (if query provided or use default)
    const searchQuery = options.searchQuery ?? 'test';
    await testSearch(manager, searchQuery);
    
    // Summary
    logSection('Summary');
    log('Tests completed. Review results above.');
    
    if (!options.headless) {
      log('\nBrowser will remain open for 10 seconds for inspection...');
      await manager.page.waitForTimeout(10000);
    }
    
  } finally {
    log('\nClosing browser...');
    await closeBrowser(manager, true);
    log('Done.');
  }
}

runTests().catch((error) => {
  console.error('\n❌ Test failed with error:', error);
  process.exit(1);
});
