#!/usr/bin/env npx tsx
/**
 * MCP Protocol Test Harness
 * 
 * Tests the MCP server by connecting a client through the actual MCP protocol,
 * rather than calling underlying functions directly. This ensures the full
 * protocol layer works correctly.
 * 
 * Usage:
 *   npm run test:mcp                       # List tools and check status
 *   npm run test:mcp -- search "query"     # Search for messages
 *   npm run test:mcp -- --json             # Output as JSON
 */

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../server.js';

interface TestOptions {
  command: 'list' | 'status' | 'search' | 'send' | 'me' | 'people' | 'favorites' | 'save' | 'unsave' | 'thread';
  query?: string;
  message?: string;
  conversationId?: string;
  messageId?: string;
  json: boolean;
  from?: number;
  size?: number;
  limit?: number;
}

function parseArgs(): TestOptions {
  const args = process.argv.slice(2);
  const options: TestOptions = {
    command: 'list',
    json: false,
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    if (arg === '--json') {
      options.json = true;
    } else if (arg === '--from' && args[i + 1]) {
      options.from = parseInt(args[i + 1], 10);
      i++;
    } else if (arg === '--size' && args[i + 1]) {
      options.size = parseInt(args[i + 1], 10);
      i++;
    } else if (arg === 'search') {
      options.command = 'search';
      // Next non-flag argument is the query
      for (let j = i + 1; j < args.length; j++) {
        if (!args[j].startsWith('--')) {
          options.query = args[j];
          break;
        }
      }
    } else if (arg === 'send') {
      options.command = 'send';
      // Next non-flag argument is the message
      for (let j = i + 1; j < args.length; j++) {
        if (!args[j].startsWith('--')) {
          options.message = args[j];
          break;
        }
      }
    } else if (arg === '--to' && args[i + 1]) {
      options.conversationId = args[i + 1];
      i++;
    } else if (arg === '--message' && args[i + 1]) {
      options.messageId = args[i + 1];
      i++;
    } else if (arg === 'me') {
      options.command = 'me';
    } else if (arg === 'status') {
      options.command = 'status';
    } else if (arg === 'list') {
      options.command = 'list';
    } else if (arg === 'people') {
      options.command = 'people';
      // Next non-flag argument is the query
      for (let j = i + 1; j < args.length; j++) {
        if (!args[j].startsWith('--')) {
          options.query = args[j];
          break;
        }
      }
    } else if (arg === 'favorites') {
      options.command = 'favorites';
    } else if (arg === 'save') {
      options.command = 'save';
    } else if (arg === 'unsave') {
      options.command = 'unsave';
    } else if (arg === 'thread') {
      options.command = 'thread';
    } else if (arg === '--limit' && args[i + 1]) {
      options.limit = parseInt(args[i + 1], 10);
      i++;
    }
  }

  return options;
}

function log(message: string): void {
  console.log(message);
}

function logSection(title: string): void {
  console.log('\n' + '─'.repeat(50));
  console.log(`  ${title}`);
  console.log('─'.repeat(50));
}

async function createTestClient(): Promise<{ client: Client; cleanup: () => Promise<void> }> {
  // Create the MCP server
  const server = await createServer();
  
  // Create linked in-memory transports
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  
  // Connect the server to its transport
  await server.connect(serverTransport);
  
  // Create and connect the client
  const client = new Client(
    { name: 'mcp-test-harness', version: '1.0.0' },
    { capabilities: {} }
  );
  
  await client.connect(clientTransport);
  
  const cleanup = async () => {
    await client.close();
    await server.close();
  };
  
  return { client, cleanup };
}

async function testListTools(client: Client, options: TestOptions): Promise<void> {
  logSection('Available Tools');
  
  const result = await client.listTools();
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  log(`Found ${result.tools.length} tools:\n`);
  
  for (const tool of result.tools) {
    log(`📦 ${tool.name}`);
    log(`   ${tool.description}`);
    
    const schema = tool.inputSchema;
    if (schema.properties) {
      const props = Object.entries(schema.properties);
      const required = new Set(schema.required ?? []);
      
      for (const [name, prop] of props) {
        const propObj = prop as { type?: string; description?: string };
        const reqMark = required.has(name) ? ' (required)' : '';
        log(`   - ${name}: ${propObj.type ?? 'any'}${reqMark}`);
        if (propObj.description) {
          log(`     ${propObj.description}`);
        }
      }
    }
    log('');
  }
}

async function testStatus(client: Client, options: TestOptions): Promise<void> {
  logSection('Teams Status');
  
  const result = await client.callTool({ name: 'teams_status', arguments: {} });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  // Parse the result
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const status = JSON.parse(textContent.text);
    
    log('\nDirect API:');
    if (status.directApi.available) {
      log(`  ✅ Available (${status.directApi.minutesRemaining} min remaining)`);
    } else {
      log('  ❌ No valid token');
    }
    
    log('\nSession:');
    log(`  Exists: ${status.session.exists ? '✅ Yes' : '❌ No'}`);
    if (status.session.likelyExpired) {
      log('  ⚠️  Likely expired');
    }
    
    log('\nBrowser:');
    log(`  Running: ${status.browser.running ? 'Yes' : 'No'}`);
  }
}

/**
 * Extracts sender name from the sender object.
 * The API returns sender as { EmailAddress: { Name: string, Address: string } }
 */
function getSenderName(sender: unknown): string | null {
  if (!sender) return null;
  if (typeof sender === 'string') return sender;
  if (typeof sender === 'object') {
    const s = sender as Record<string, unknown>;
    // Handle { EmailAddress: { Name: string } } structure
    if (s.EmailAddress && typeof s.EmailAddress === 'object') {
      const email = s.EmailAddress as Record<string, unknown>;
      if (email.Name) return String(email.Name);
      if (email.Address) return String(email.Address);
    }
    // Handle { name: string } structure
    if (s.name) return String(s.name);
    if (s.Name) return String(s.Name);
  }
  return null;
}

async function testSearch(client: Client, options: TestOptions): Promise<void> {
  if (!options.query) {
    console.error('❌ Error: Search query required');
    console.error('   Usage: npm run test:mcp -- search "your query"');
    process.exit(1);
  }
  
  if (!options.json) {
    logSection(`Search: "${options.query}"`);
  }
  
  const args: Record<string, unknown> = { query: options.query };
  if (options.from !== undefined) args.from = options.from;
  if (options.size !== undefined) args.size = options.size;
  
  if (!options.json) {
    log(`Calling teams_search via MCP protocol...`);
    log(`Arguments: ${JSON.stringify(args)}\n`);
  }
  
  const result = await client.callTool({ name: 'teams_search', arguments: args });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  // Parse the result
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Search failed: ${response.error}`);
      return;
    }
    
    log(`✅ Search successful (mode: ${response.mode})`);
    log(`   Results: ${response.resultCount}`);
    
    if (response.pagination) {
      const p = response.pagination;
      log(`   Pagination: from=${p.from}, size=${p.size}, returned=${p.returned}`);
      if (p.total !== undefined) {
        log(`   Total available: ${p.total}`);
      }
      if (p.hasMore) {
        log(`   More results available (use --from ${p.nextFrom})`);
      }
    }
    
    if (response.results && response.results.length > 0) {
      log('\n📋 Results:\n');
      
      for (let i = 0; i < response.results.length; i++) {
        const r = response.results[i];
        const num = (response.pagination?.from ?? 0) + i + 1;
        const preview = (r.content ?? '').substring(0, 100).replace(/\n/g, ' ');
        
        log(`${num}. ${preview}${r.content?.length > 100 ? '...' : ''}`);
        
        const senderName = getSenderName(r.sender);
        if (senderName) log(`   From: ${senderName}`);
        if (r.teamName) log(`   Team: ${r.teamName}`);
        if (r.channelName && r.channelName !== r.teamName) log(`   Channel: ${r.channelName}`);
        if (r.timestamp) log(`   Time: ${r.timestamp}`);
        log('');
      }
    }
  }
}

async function testSend(client: Client, options: TestOptions): Promise<void> {
  if (!options.message) {
    console.error('❌ Error: Message content required');
    console.error('   Usage: npm run test:mcp -- send "your message"');
    process.exit(1);
  }
  
  if (!options.json) {
    logSection(`Send Message`);
  }
  
  const args: Record<string, unknown> = { content: options.message };
  if (options.conversationId) args.conversationId = options.conversationId;
  
  if (!options.json) {
    log(`Calling teams_send_message via MCP protocol...`);
    log(`Arguments: ${JSON.stringify(args)}\n`);
  }
  
  const result = await client.callTool({ name: 'teams_send_message', arguments: args });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  // Parse the result
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Send failed: ${response.error}`);
      return;
    }
    
    log(`✅ Message sent successfully!`);
    log(`   Message ID: ${response.messageId}`);
    if (response.timestamp) {
      log(`   Timestamp: ${new Date(response.timestamp).toISOString()}`);
    }
    log(`   Conversation: ${response.conversationId}`);
  }
}

async function testMe(client: Client, options: TestOptions): Promise<void> {
  if (!options.json) {
    logSection(`Get Current User`);
  }
  
  if (!options.json) {
    log(`Calling teams_get_me via MCP protocol...`);
  }
  
  const result = await client.callTool({ name: 'teams_get_me', arguments: {} });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  // Parse the result
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Failed: ${response.error}`);
      return;
    }
    
    const p = response.profile;
    log(`\n👤 Current User\n`);
    log(`   Name: ${p.displayName}`);
    log(`   Email: ${p.email}`);
    log(`   ID: ${p.id}`);
    log(`   MRI: ${p.mri}`);
    if (p.tenantId) {
      log(`   Tenant: ${p.tenantId}`);
    }
  }
}

async function testPeople(client: Client, options: TestOptions): Promise<void> {
  if (!options.query) {
    console.error('❌ Error: Search query required');
    console.error('   Usage: npm run test:mcp -- people "name or email"');
    process.exit(1);
  }
  
  if (!options.json) {
    logSection(`Search People: "${options.query}"`);
    log(`Calling teams_search_people via MCP protocol...`);
  }
  
  const result = await client.callTool({ 
    name: 'teams_search_people', 
    arguments: { query: options.query } 
  });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Failed: ${response.error}`);
      return;
    }
    
    log(`\n✅ Found ${response.returned} people:\n`);
    
    for (const p of response.results) {
      log(`👤 ${p.displayName}`);
      if (p.email) log(`   Email: ${p.email}`);
      if (p.jobTitle) log(`   Title: ${p.jobTitle}`);
      if (p.department) log(`   Dept: ${p.department}`);
      log(`   MRI: ${p.mri}`);
      log('');
    }
  }
}

async function testFavorites(client: Client, options: TestOptions): Promise<void> {
  if (!options.json) {
    logSection(`Get Favourites`);
    log(`Calling teams_get_favorites via MCP protocol...`);
  }
  
  const result = await client.callTool({ name: 'teams_get_favorites', arguments: {} });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Failed: ${response.error}`);
      return;
    }
    
    log(`\n⭐ Found ${response.count} favourites:\n`);
    
    for (const f of response.favorites) {
      log(`   ${f.conversationId}`);
    }
  }
}

async function testThread(client: Client, options: TestOptions): Promise<void> {
  if (!options.conversationId) {
    console.error(`❌ Error: Conversation ID required`);
    console.error(`   Usage: npm run test:mcp -- thread --to "conversationId" [--limit 50]`);
    process.exit(1);
  }
  
  if (!options.json) {
    logSection(`Get Thread Messages`);
    log(`Calling teams_get_thread via MCP protocol...`);
  }
  
  const args: Record<string, unknown> = { conversationId: options.conversationId };
  if (options.limit) args.limit = options.limit;
  
  const result = await client.callTool({ name: 'teams_get_thread', arguments: args });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Failed: ${response.error}`);
      return;
    }
    
    log(`\n✅ Got ${response.messageCount} messages from thread:\n`);
    
    for (const msg of response.messages || []) {
      const preview = msg.content.substring(0, 100).replace(/\n/g, ' ');
      const sender = msg.sender?.displayName || msg.sender?.mri || 'Unknown';
      const time = new Date(msg.timestamp).toLocaleString();
      const fromMe = msg.isFromMe ? ' (you)' : '';
      
      log(`📝 ${sender}${fromMe} - ${time}`);
      log(`   ${preview}${msg.content.length > 100 ? '...' : ''}`);
      log('');
    }
  }
}

async function testSaveMessage(client: Client, options: TestOptions, save: boolean): Promise<void> {
  if (!options.conversationId || !options.messageId) {
    console.error(`❌ Error: Conversation ID and message ID required`);
    console.error(`   Usage: npm run test:mcp -- ${save ? 'save' : 'unsave'} --to "conversationId" --message "messageId"`);
    process.exit(1);
  }
  
  if (!options.json) {
    logSection(save ? `Save Message` : `Unsave Message`);
    log(`Calling teams_${save ? 'save' : 'unsave'}_message via MCP protocol...`);
  }
  
  const result = await client.callTool({ 
    name: save ? 'teams_save_message' : 'teams_unsave_message', 
    arguments: { 
      conversationId: options.conversationId,
      messageId: options.messageId,
    } 
  });
  
  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');
  
  if (textContent?.text) {
    const response = JSON.parse(textContent.text);
    
    if (!response.success) {
      log(`❌ Failed: ${response.error}`);
      return;
    }
    
    log(`\n✅ Message ${save ? 'saved' : 'unsaved'} successfully!`);
    log(`   Conversation: ${response.conversationId}`);
    log(`   Message: ${response.messageId}`);
  }
}

async function main(): Promise<void> {
  const options = parseArgs();
  
  if (!options.json) {
    console.log('\n🧪 MCP Protocol Test Harness');
    console.log('============================');
  }
  
  let cleanup: (() => Promise<void>) | null = null;
  
  try {
    const { client, cleanup: cleanupFn } = await createTestClient();
    cleanup = cleanupFn;
    
    if (!options.json) {
      log('\n✅ Connected to MCP server via in-memory transport');
    }
    
    switch (options.command) {
      case 'list':
        await testListTools(client, options);
        break;
      case 'status':
        await testStatus(client, options);
        break;
      case 'search':
        await testSearch(client, options);
        break;
      case 'send':
        await testSend(client, options);
        break;
      case 'me':
        await testMe(client, options);
        break;
      case 'people':
        await testPeople(client, options);
        break;
      case 'favorites':
        await testFavorites(client, options);
        break;
      case 'save':
        await testSaveMessage(client, options, true);
        break;
      case 'unsave':
        await testSaveMessage(client, options, false);
        break;
      case 'thread':
        await testThread(client, options);
        break;
    }
    
    if (!options.json) {
      logSection('Complete');
      log('MCP protocol test finished successfully.');
    }
    
  } catch (error) {
    console.error('\n❌ Error:', error instanceof Error ? error.message : String(error));
    if (error instanceof Error && error.stack) {
      console.error('\nStack trace:');
      console.error(error.stack);
    }
    process.exit(1);
  } finally {
    if (cleanup) {
      await cleanup();
    }
  }
}

main();
