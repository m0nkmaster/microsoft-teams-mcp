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
import { searchTeamsWithPagination } from '../teams/search.js';
import {
  hasSessionState,
  getSessionAge,
  clearSessionState,
} from '../auth/session-store.js';
import {
  hasValidSubstrateToken,
  getSubstrateTokenStatus,
  extractMessageAuth,
  getUserProfile,
  clearTokenCache,
} from '../auth/token-extractor.js';
import { searchMessages } from '../api/substrate-api.js';
import { sendMessage, sendNoteToSelf } from '../api/chatsvc-api.js';

type Command = 'status' | 'search' | 'login' | 'send' | 'me' | 'help';

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
  search <query>      Search Teams for messages
  send <message>      Send a message to yourself (notes)
  send --to <id>      Send a message to a specific conversation
  me                  Get current user profile (email, name, Teams ID)
  login               Log in to Teams (opens browser)
  login --force       Force new login (clears existing session)
  help                Show this help message

Options:
  --headless          Run browser in headless mode (where applicable)
  --json              Output results as JSON
  --debug             Show debug output for troubleshooting
  --browser           Force browser mode (skip direct API)

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
  npm run cli -- search "query" --browser
  npm run cli -- send "Test message to myself"
  npm run cli -- login --force
`);
}

async function commandStatus(flags: Set<string>): Promise<void> {
  const hasSession = hasSessionState();
  const sessionAge = getSessionAge();
  const tokenStatus = getSubstrateTokenStatus();

  const status = {
    directApi: {
      available: tokenStatus.hasToken,
      expiresAt: tokenStatus.expiresAt,
      minutesRemaining: tokenStatus.minutesRemaining,
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
    console.log('\n📊 Status\n');

    if (status.directApi.available) {
      console.log(`Direct API: ✅ Available (${status.directApi.minutesRemaining} min remaining)`);
    } else {
      console.log('Direct API: ❌ No valid token (browser login required)');
    }

    console.log(`Session exists: ${status.session.exists ? '✅ Yes' : '❌ No'}`);
    if (status.session.ageHours !== null) {
      console.log(`Session age: ${status.session.ageHours} hours`);
      if (status.session.likelyExpired) {
        console.log('⚠️  Session may be expired');
      }
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

  const headless = flags.has('headless');
  const asJson = flags.has('json');
  const debug = flags.has('debug');
  const forceBrowser = flags.has('browser');

  const from = options.has('from') ? parseInt(options.get('from')!, 10) : 0;
  const size = options.has('size') ? parseInt(options.get('size')!, 10) : 25;
  const maxResults = options.has('maxResults') ? parseInt(options.get('maxResults')!, 10) : 25;

  if (!asJson) {
    console.log(`\n🔍 Searching for: "${query}"`);
    if (from > 0) {
      console.log(`   Offset: ${from}, Size: ${size}`);
    }
  }

  // Try direct API first
  if (!forceBrowser && hasValidSubstrateToken()) {
    if (!asJson) {
      console.log('   Using direct API...\n');
    }

    const result = await searchMessages(query, { from, size, maxResults });

    if (result.ok) {
      if (asJson) {
        console.log(JSON.stringify({
          mode: 'direct-api',
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
      return;
    }

    if (!asJson) {
      console.log(`   Direct API failed: ${result.error.message}`);
      console.log('   Falling back to browser...\n');
    }
  } else if (!asJson && !forceBrowser) {
    console.log('   No valid token, using browser...\n');
  } else if (!asJson) {
    console.log('   Using browser (--browser flag)...\n');
  }

  // Fall back to browser-based search
  let manager: BrowserManager | null = null;

  try {
    manager = await createBrowserContext({ headless });

    await ensureAuthenticated(
      manager.page,
      manager.context,
      asJson ? undefined : (msg) => console.log(`   ${msg}`)
    );

    const { results, pagination } = await searchTeamsWithPagination(manager.page, query, {
      maxResults,
      from,
      size,
      waitMs: 10000,
      debug,
    });

    await manager.page.waitForTimeout(3000);

    if (asJson) {
      console.log(JSON.stringify({
        mode: 'browser',
        query,
        count: results.length,
        pagination: {
          from: pagination.from,
          size: pagination.size,
          returned: pagination.returned,
          total: pagination.total,
          hasMore: pagination.hasMore,
          nextFrom: pagination.hasMore ? pagination.from + pagination.returned : undefined,
        },
        results,
      }, null, 2));
    } else {
      printResults(results, pagination);
      console.log('💡 Session saved. Future searches can use direct API.');
    }
  } finally {
    if (manager) {
      await closeBrowser(manager, true);
    }
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
