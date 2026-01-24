/**
 * Search-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { hasValidSubstrateToken } from '../auth/token-extractor.js';
import { searchMessages, searchChannels } from '../api/substrate-api.js';
import { getThreadMessages } from '../api/chatsvc-api.js';
import { searchTeamsWithPagination } from '../teams/search.js';
import { closeBrowser } from '../browser/context.js';
import {
  DEFAULT_PAGE_SIZE,
  MAX_PAGE_SIZE,
  DEFAULT_THREAD_LIMIT,
  MAX_THREAD_LIMIT,
  DEFAULT_CHANNEL_LIMIT,
  MAX_CHANNEL_LIMIT,
  MSAL_TOKEN_DELAY_MS,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SearchInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  maxResults: z.number().optional().default(DEFAULT_PAGE_SIZE),
  from: z.number().min(0).optional().default(0),
  size: z.number().min(1).max(MAX_PAGE_SIZE).optional().default(DEFAULT_PAGE_SIZE),
});

export const GetThreadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional().default(DEFAULT_THREAD_LIMIT),
});

export const FindChannelInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  limit: z.number().min(1).max(MAX_CHANNEL_LIMIT).optional().default(DEFAULT_CHANNEL_LIMIT),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const searchToolDefinition: Tool = {
  name: 'teams_search',
  description: 'Search for messages in Microsoft Teams. Returns matching messages with sender, timestamp, content, conversationId (for replies), and pagination info. Supports search operators: from:email, sent:today/lastweek, in:channel, hasattachment:true, "Name" for @mentions. Combine with NOT to exclude (e.g., NOT from:rob@co.com).',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search query with optional operators. Examples: "budget report", "from:sarah@co.com sent:lastweek", "\"Rob Smith\" NOT from:rob@co.com" (find @mentions of Rob). IMPORTANT: "@me", "from:me", "to:me" do NOT work - use teams_get_me first to get actual email/displayName, then use those values.',
      },
      maxResults: {
        type: 'number',
        description: 'Maximum number of results to return (default: 25)',
      },
      from: {
        type: 'number',
        description: 'Starting offset for pagination (0-based, default: 0). Use this to get subsequent pages of results.',
      },
      size: {
        type: 'number',
        description: 'Page size (default: 25). Number of results per page.',
      },
    },
    required: ['query'],
  },
};

const getThreadToolDefinition: Tool = {
  name: 'teams_get_thread',
  description: 'Get messages from a Teams conversation/thread. Use this to see replies to a message, check thread context, or read recent messages in a chat. Requires a conversationId (available from search results).',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get messages from (e.g., "19:abc@thread.tacv2" from search results)',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of messages to return (default: 50, max: 200)',
      },
    },
    required: ['conversationId'],
  },
};

const findChannelToolDefinition: Tool = {
  name: 'teams_find_channel',
  description: 'Find Teams channels by name. Searches both (1) channels in teams you\'re a member of (reliable) and (2) channels across the organisation (discovery). Results indicate whether you\'re already a member via the isMember field. Channel IDs can be used with teams_get_thread to read messages.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Channel name to search for (partial match)',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of results (default: 10, max: 50)',
      },
    },
    required: ['query'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleSearch(
  input: z.infer<typeof SearchInputSchema>,
  ctx: ToolContext
): Promise<ToolResult> {
  // Try direct API first
  if (hasValidSubstrateToken()) {
    const result = await searchMessages(input.query, {
      maxResults: input.maxResults,
      from: input.from,
      size: input.size,
    });

    if (result.ok) {
      return {
        success: true,
        data: {
          mode: 'direct-api',
          query: input.query,
          resultCount: result.value.results.length,
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
        },
      };
    }

    // Log error but fall through to browser
    console.error('[search] Direct API failed:', result.error.message);
  }

  // Fall back to browser-based search
  const manager = await ctx.server.ensureBrowser(false);

  const { results, pagination } = await searchTeamsWithPagination(
    manager.page,
    input.query,
    {
      maxResults: input.maxResults,
      from: input.from,
      size: input.size,
    }
  );

  // Wait for MSAL to store tokens
  await manager.page.waitForTimeout(MSAL_TOKEN_DELAY_MS);

  // Close browser after search
  await closeBrowser(manager, true);
  ctx.server.resetBrowserState();

  return {
    success: true,
    data: {
      mode: 'browser',
      note: 'Session saved. Future searches will use direct API.',
      query: input.query,
      resultCount: results.length,
      pagination: {
        from: pagination.from,
        size: pagination.size,
        returned: pagination.returned,
        total: pagination.total,
        hasMore: pagination.hasMore,
        nextFrom: pagination.hasMore
          ? pagination.from + pagination.returned
          : undefined,
      },
      results,
    },
  };
}

async function handleGetThread(
  input: z.infer<typeof GetThreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getThreadMessages(input.conversationId, { limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      messageCount: result.value.messages.length,
      messages: result.value.messages,
    },
  };
}

async function handleFindChannel(
  input: z.infer<typeof FindChannelInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await searchChannels(input.query, input.limit);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      query: input.query,
      count: result.value.returned,
      channels: result.value.results,
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const searchTool: RegisteredTool<typeof SearchInputSchema> = {
  definition: searchToolDefinition,
  schema: SearchInputSchema,
  handler: handleSearch,
};

export const getThreadTool: RegisteredTool<typeof GetThreadInputSchema> = {
  definition: getThreadToolDefinition,
  schema: GetThreadInputSchema,
  handler: handleGetThread,
};

export const findChannelTool: RegisteredTool<typeof FindChannelInputSchema> = {
  definition: findChannelToolDefinition,
  schema: FindChannelInputSchema,
  handler: handleFindChannel,
};

/** All search-related tools. */
export const searchTools = [searchTool, getThreadTool, findChannelTool];
