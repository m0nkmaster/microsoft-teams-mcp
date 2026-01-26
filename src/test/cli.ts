#!/usr/bin/env npx tsx
/**
 * CLI tool to interact with Teams MCP functionality directly.
 * Useful for testing individual operations.
 *
 * Usage:
 *   npm run cli -- status
 *   npm run cli -- search "your query"
 *   npm run cli -- login
 *   npm run cli -- login --force
 */

import { createBrowserContext, closeBrowser, type BrowserManager } from '../browser/context.js';
import { ensureAuthenticated, forceNewLogin } from '../browser/auth.js';
import {
  hasSessionState,
  getSessionAge,
  clearSessionState,
} from '../auth/session-store.js';
import {
  hasValidSubstrateToken,
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  getUserProfile,
  clearTokenCache,
} from '../auth/token-extractor.js';
import { refreshTokensViaBrowser } from '../auth/token-refresh.js';
import { searchMessages } from '../api/substrate-api.js';
import { sendMessage, sendNoteToSelf } from '../api/chatsvc-api.js';

type Command = 'status' | 'search' | 'login' | 'send' | 'me' | 'refresh' | 'refresh-debug' | 'help';

interface CliArgs {
  command: Command;
  args: string[];
  flags: Set<string>;
  options: Map<string, string>;
}

function parseArgs(): CliArgs {
  const args = process.argv.slice(2);
  const command = (args[0] ?? 'help') as Command;
  const flags = new Set<string>();
  const options = new Map<string, string>();
  const remainingArgs: string[] = [];

  for (let i = 1; i < args.length; i++) {
    const arg = args[i];
    if (arg.startsWith('--') && arg.includes('=')) {
      const [key, value] = arg.slice(2).split('=', 2);
      options.set(key, value);
    } else if (arg.startsWith('--')) {
      const key = arg.slice(2);
      const next = args[i + 1];
      if (next && !next.startsWith('-')) {
        if (/^\d+$/.test(next)) {
          options.set(key, next);
          i++;
        } else {
          flags.add(key);
        }
      } else {
        flags.add(key);
      }
    } else if (arg.startsWith('-')) {
      flags.add(arg.slice(1));
    } else {
      remainingArgs.push(arg);
    }
  }

  return { command, args: remainingArgs, flags, options };
}

function printHelp(): void {
  console.log(`
Teams MCP CLI

Commands:
  status              Check session and authentication status
  search <query>      Search Teams for messages (requires valid token)
  send <message>      Send a message to yourself (notes)
  send --to <id>      Send a message to a specific conversation
  me                  Get current user profile (email, name, Teams ID)
  login               Log in to Teams (opens browser)
  login --force       Force new login (clears existing session)
  refresh             Test OAuth token refresh (shows before/after status)
  refresh-debug       Debug token refresh with visible browser
  help                Show this help message

Options:
  --json              Output results as JSON

Pagination Options (for search):
  --from <n>          Starting offset (default: 0, for page 2 use --from 25)
  --size <n>          Page size (default: 25)
  --maxResults <n>    Maximum results to return (default: 25)

Send Options:
  --to <conversationId>   Send to a specific conversation (default: 48:notes = self)

Examples:
  npm run cli -- status
  npm run cli -- search "meeting notes"
  npm run cli -- search "project update" --json
  npm run cli -- search "query" --from 25
  npm run cli -- send "Test message to myself"
  npm run cli -- login --force
  npm run cli -- refresh
`);
}

async function commandStatus(flags: Set<string>): Promise<void> {
  const hasSession = hasSessionState();
  const sessionAge = getSessionAge();
  const substrateStatus = getSubstrateTokenStatus();
  const messagingStatus = getMessageAuthStatus();

  // Check if refresh will trigger soon (within 10 minutes)
  const REFRESH_THRESHOLD_MINS = 10;
  const willRefreshSoon = substrateStatus.minutesRemaining !== undefined && 
    substrateStatus.minutesRemaining <= REFRESH_THRESHOLD_MINS;

  const status = {
    search: {
      available: substrateStatus.hasToken,
      expiresAt: substrateStatus.expiresAt,
      minutesRemaining: substrateStatus.minutesRemaining,
      willAutoRefresh: willRefreshSoon,
    },
    messaging: {
      available: messagingStatus.hasToken,
      expiresAt: messagingStatus.expiresAt,
      minutesRemaining: messagingStatus.minutesRemaining,
    },
    session: {
      exists: hasSession,
      ageHours: sessionAge !== null ? Math.round(sessionAge * 10) / 10 : null,
      likelyExpired: sessionAge !== null ? sessionAge > 12 : null,
    },
  };

  if (flags.has('json')) {
    console.log(JSON.stringify(status, null, 2));
  } else {
    console.log('\n📊 Token Status\n');

    // Search token (Substrate)
    console.log('Search API (Substrate):');
    if (status.search.available) {
      console.log(`   Status: ✅ Valid`);
      console.log(`   Expires: ${status.search.expiresAt}`);
      console.log(`   Remaining: ${status.search.minutesRemaining} minutes`);
      if (status.search.willAutoRefresh) {
        console.log(`   ⚡ Auto-refresh will trigger (< ${REFRESH_THRESHOLD_MINS} min remaining)`);
      }
    } else {
      console.log('   Status: ❌ No valid token');
    }

    // Messaging token (Skype)
    console.log('\nMessaging API (Skype):');
    if (status.messaging.available) {
      console.log('   Status: ✅ Valid');
      if (status.messaging.expiresAt) {
        console.log(`   Expires: ${status.messaging.expiresAt}`);
        console.log(`   Remaining: ${status.messaging.minutesRemaining} minutes`);
      }
    } else {
      console.log('   Status: ❌ No valid token');
    }

    // Session info
    console.log('\nSession:');
    console.log(`   Exists: ${status.session.exists ? '✅ Yes' : '❌ No'}`);
    if (status.session.ageHours !== null) {
      console.log(`   Age: ${status.session.ageHours} hours`);
      if (status.session.likelyExpired) {
        console.log('   ⚠️  Session may be expired (>12 hours old)');
      }
    }

    // Action hint
    if (!status.search.available || !status.messaging.available) {
      console.log('\n💡 Run: npm run cli -- login');
    } else if (status.search.willAutoRefresh) {
      console.log('\n💡 Run: npm run cli -- refresh (to test auto-refresh)');
    }
  }
}

async function commandSearch(
  query: string,
  flags: Set<string>,
  options: Map<string, string>
): Promise<void> {
  if (!query) {
    console.error('❌ Error: Search query required');
    console.error('   Usage: npm run cli -- search "your query"');
    process.exit(1);
  }

  const asJson = flags.has('json');

  const from = options.has('from') ? parseInt(options.get('from')!, 10) : 0;
  const size = options.has('size') ? parseInt(options.get('size')!, 10) : 25;
  const maxResults = options.has('maxResults') ? parseInt(options.get('maxResults')!, 10) : 25;

  if (!hasValidSubstrateToken()) {
    if (asJson) {
      console.log(JSON.stringify({ success: false, error: 'No valid token. Please run: npm run cli -- login' }, null, 2));
    } else {
      console.error('❌ No valid token. Please run: npm run cli -- login');
    }
    process.exit(1);
  }

  if (!asJson) {
    console.log(`\n🔍 Searching for: "${query}"`);
    if (from > 0) {
      console.log(`   Offset: ${from}, Size: ${size}`);
    }
  }

  const result = await searchMessages(query, { from, size, maxResults });

  if (!result.ok) {
    if (asJson) {
      console.log(JSON.stringify({ success: false, error: result.error.message }, null, 2));
    } else {
      console.error(`❌ Search failed: ${result.error.message}`);
    }
    process.exit(1);
  }

  if (asJson) {
    console.log(JSON.stringify({
      query,
      count: result.value.results.length,
      pagination: {
        from: result.value.pagination.from,
        size: result.value.pagination.size,
        returned: result.value.pagination.returned,
        total: result.value.pagination.total,
        hasMore: result.value.pagination.hasMore,
        nextFrom: result.value.pagination.hasMore
          ? result.value.pagination.from + result.value.pagination.returned
          : undefined,
      },
      results: result.value.results,
    }, null, 2));
  } else {
    printResults(result.value.results, result.value.pagination);
  }
}

function printResults(
  results: import('../types/teams.js').TeamsSearchResult[],
  pagination: import('../types/teams.js').SearchPaginationResult
): void {
  console.log(`\n📋 Found ${results.length} results`);
  if (pagination.total !== undefined) {
    console.log(`   Total available: ${pagination.total}`);
  }
  if (pagination.hasMore) {
    console.log(`   More results available (use --from ${pagination.from + pagination.returned})`);
  }
  console.log();

  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    console.log(`${pagination.from + i + 1}. ${r.content.substring(0, 100).replace(/\n/g, ' ')}${r.content.length > 100 ? '...' : ''}`);
    if (r.sender) console.log(`   From: ${r.sender}`);
    if (r.timestamp) console.log(`   Time: ${r.timestamp}`);
    console.log();
  }
}

async function commandSend(
  message: string,
  flags: Set<string>,
  options: Map<string, string>
): Promise<void> {
  if (!message) {
    console.error('❌ Error: Message content required');
    console.error('   Usage: npm run cli -- send "your message"');
    process.exit(1);
  }

  const asJson = flags.has('json');
  const conversationId = options.get('to') || '48:notes';

  const auth = extractMessageAuth();
  if (!auth) {
    console.error('❌ No valid authentication. Please run: npm run cli -- login');
    process.exit(1);
  }

  if (!asJson) {
    if (conversationId === '48:notes') {
      console.log(`\n📝 Sending note to yourself...`);
    } else {
      console.log(`\n📤 Sending message to: ${conversationId}`);
    }
    console.log(`   Content: "${message.substring(0, 50)}${message.length > 50 ? '...' : ''}"`);
  }

  const result = conversationId === '48:notes'
    ? await sendNoteToSelf(message)
    : await sendMessage(conversationId, message);

  if (asJson) {
    console.log(JSON.stringify(
      result.ok
        ? { success: true, ...result.value }
        : { success: false, error: result.error.message },
      null,
      2
    ));
  } else {
    if (result.ok) {
      console.log('\n✅ Message sent successfully!');
      console.log(`   Message ID: ${result.value.messageId}`);
      if (result.value.timestamp) {
        console.log(`   Timestamp: ${new Date(result.value.timestamp).toISOString()}`);
      }
    } else {
      console.error(`\n❌ Failed to send message: ${result.error.message}`);
      process.exit(1);
    }
  }
}

async function commandMe(flags: Set<string>): Promise<void> {
  const asJson = flags.has('json');

  const profile = getUserProfile();

  if (!profile) {
    if (asJson) {
      console.log(JSON.stringify({ success: false, error: 'No valid session' }, null, 2));
    } else {
      console.error('❌ No valid session. Please run: npm run cli -- login');
    }
    process.exit(1);
  }

  if (asJson) {
    console.log(JSON.stringify({ success: true, profile }, null, 2));
  } else {
    console.log('\n👤 Current User\n');
    console.log(`   Name: ${profile.displayName}`);
    console.log(`   Email: ${profile.email}`);
    console.log(`   ID: ${profile.id}`);
    console.log(`   MRI: ${profile.mri}`);
    if (profile.tenantId) {
      console.log(`   Tenant: ${profile.tenantId}`);
    }
  }
}

async function commandRefresh(flags: Set<string>): Promise<void> {
  const asJson = flags.has('json');

  // Show current status
  const beforeStatus = getSubstrateTokenStatus();

  if (!asJson) {
    console.log('\n🔄 Token Refresh Test\n');
    console.log('Before refresh:');
    if (beforeStatus.hasToken) {
      console.log(`   Token valid: ✅ Yes`);
      console.log(`   Expires at: ${beforeStatus.expiresAt}`);
      console.log(`   Minutes remaining: ${beforeStatus.minutesRemaining}`);
    } else {
      console.log(`   Token valid: ❌ No`);
    }

    console.log('\nOpening headless browser to refresh tokens...');
  }

  // Attempt refresh via headless browser
  const result = await refreshTokensViaBrowser();

  if (!result.ok) {
    if (asJson) {
      console.log(JSON.stringify({
        success: false,
        error: result.error.message,
        code: result.error.code,
        before: beforeStatus,
      }, null, 2));
    } else {
      console.log(`\n❌ Refresh failed: ${result.error.message}`);
      if (result.error.code) {
        console.log(`   Error code: ${result.error.code}`);
      }
      if (result.error.suggestions) {
        console.log(`   Suggestions: ${result.error.suggestions.join(', ')}`);
      }
    }
    process.exit(1);
  }

  // Show new status
  const afterStatus = getSubstrateTokenStatus();

  if (asJson) {
    console.log(JSON.stringify({
      success: true,
      before: beforeStatus,
      after: afterStatus,
      refreshResult: {
        previousExpiry: result.value.previousExpiry.toISOString(),
        newExpiry: result.value.newExpiry.toISOString(),
        minutesGained: result.value.minutesGained,
        refreshNeeded: result.value.refreshNeeded,
      },
    }, null, 2));
  } else {
    if (result.value.refreshNeeded) {
      console.log('\n✅ Token refreshed!\n');
      console.log('After refresh:');
      console.log(`   Token valid: ✅ Yes`);
      console.log(`   Expires at: ${afterStatus.expiresAt}`);
      console.log(`   Minutes remaining: ${afterStatus.minutesRemaining}`);
      console.log(`   Time gained: +${result.value.minutesGained} minutes`);
    } else {
      console.log('\n✅ Token is still valid (no refresh needed)\n');
      console.log('Current token:');
      console.log(`   Expires at: ${afterStatus.expiresAt}`);
      console.log(`   Minutes remaining: ${afterStatus.minutesRemaining}`);
      console.log('\n   Note: MSAL only refreshes tokens when they\'re close to expiry.');
      console.log('   Proactive refresh will trigger when <10 minutes remain.');
    }
  }
}

async function commandRefreshDebug(flags: Set<string>): Promise<void> {
  const asJson = flags.has('json');
  
  // Show current status
  const beforeStatus = getSubstrateTokenStatus();

  if (!asJson) {
    console.log('\n🔍 Token Refresh Debug (Visible Browser)\n');
    console.log('Before:');
    console.log(`   Token valid: ${beforeStatus.hasToken ? '✅ Yes' : '❌ No'}`);
    console.log(`   Expires at: ${beforeStatus.expiresAt}`);
    console.log(`   Minutes remaining: ${beforeStatus.minutesRemaining}`);
    console.log('\nOpening VISIBLE browser to debug token refresh...');
  }

  // Import browser and auth functions
  const { createBrowserContext, closeBrowser } = await import('../browser/context.js');
  const { ensureAuthenticated } = await import('../browser/auth.js');

  let manager: Awaited<ReturnType<typeof createBrowserContext>> | null = null;

  try {
    // Open VISIBLE browser
    manager = await createBrowserContext({ headless: false });

    // Use the same auth flow that works for login
    await ensureAuthenticated(manager.page, manager.context, (msg) => {
      if (!asJson) {
        console.log(`   ${msg}`);
      }
    });

    // Close browser (session already saved by ensureAuthenticated)
    await closeBrowser(manager, false);
    manager = null;

    // Clear cache and check new token
    clearTokenCache();
    const afterStatus = getSubstrateTokenStatus();

    if (asJson) {
      console.log(JSON.stringify({
        before: beforeStatus,
        after: afterStatus,
        refreshed: afterStatus.hasToken && afterStatus.expiresAt !== beforeStatus.expiresAt,
      }, null, 2));
    } else {
      console.log('\nAfter:');
      console.log(`   Token valid: ${afterStatus.hasToken ? '✅ Yes' : '❌ No'}`);
      console.log(`   Expires at: ${afterStatus.expiresAt}`);
      console.log(`   Minutes remaining: ${afterStatus.minutesRemaining}`);

      if (afterStatus.hasToken && afterStatus.expiresAt !== beforeStatus.expiresAt) {
        console.log('\n✅ Token was refreshed!');
      } else if (afterStatus.hasToken) {
        console.log('\n⚠️  Token is valid but was not refreshed (same expiry)');
      } else {
        console.log('\n❌ Token still invalid after browser session');
      }
    }

  } catch (error) {
    // Clean up browser if still open
    if (manager) {
      try {
        await closeBrowser(manager, false);
      } catch {
        // Ignore cleanup errors
      }
    }
    throw error;
  }
}

async function commandLogin(flags: Set<string>): Promise<void> {
  const force = flags.has('force');

  if (force) {
    console.log('🔄 Forcing new login (clearing existing session)...');
    clearSessionState();
    clearTokenCache();
  } else {
    console.log('🔐 Starting login flow...');
  }

  let manager: BrowserManager | null = null;

  try {
    manager = await createBrowserContext({ headless: false });

    if (force) {
      await forceNewLogin(
        manager.page,
        manager.context,
        (msg) => console.log(`   ${msg}`)
      );
    } else {
      await ensureAuthenticated(
        manager.page,
        manager.context,
        (msg) => console.log(`   ${msg}`)
      );
    }

    console.log('\n✅ Login successful! Session has been saved.');
    console.log('   You can now run searches in headless mode.');

  } finally {
    if (manager) {
      await closeBrowser(manager, true);
    }
  }
}

async function main(): Promise<void> {
  const { command, args, flags, options } = parseArgs();

  try {
    switch (command) {
      case 'status':
        await commandStatus(flags);
        break;

      case 'search':
        await commandSearch(args.join(' '), flags, options);
        break;

      case 'send':
        await commandSend(args.join(' '), flags, options);
        break;

      case 'me':
        await commandMe(flags);
        break;

      case 'refresh':
        await commandRefresh(flags);
        break;

      case 'refresh-debug':
        await commandRefreshDebug(flags);
        break;

      case 'login':
        await commandLogin(flags);
        break;

      case 'help':
      default:
        printHelp();
        break;
    }
  } catch (error) {
    console.error('\n❌ Error:', error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}

main();
