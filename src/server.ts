/**
 * MCP Server implementation for Teams search.
 * Exposes tools and resources for interacting with Microsoft Teams.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  type Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

import { createBrowserContext, closeBrowser, type BrowserManager } from './browser/context.js';
import { ensureAuthenticated, getAuthStatus, forceNewLogin } from './browser/auth.js';
import { searchTeamsWithPagination } from './teams/search.js';

// Auth modules
import {
  hasSessionState,
  isSessionLikelyExpired,
  clearSessionState,
} from './auth/session-store.js';
import {
  hasValidSubstrateToken,
  getSubstrateTokenStatus,
  extractMessageAuth,
  extractCsaToken,
  getUserProfile,
  clearTokenCache,
} from './auth/token-extractor.js';

// API modules
import { searchMessages, searchPeople, getFrequentContacts, searchChannels } from './api/substrate-api.js';
import { sendMessage, sendNoteToSelf, replyToThread, getThreadMessages, saveMessage, unsaveMessage, getOneOnOneChatId } from './api/chatsvc-api.js';
import { getFavorites, addFavorite, removeFavorite } from './api/csa-api.js';

// Types
import { ErrorCode, createError, type McpError } from './types/errors.js';

// Tool definitions
const TOOLS: Tool[] = [
  {
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
  },
  {
    name: 'teams_login',
    description: 'Trigger manual login flow for Microsoft Teams. Use this if the session has expired or you need to switch accounts.',
    inputSchema: {
      type: 'object',
      properties: {
        forceNew: {
          type: 'boolean',
          description: 'Force a new login even if a session exists (default: false)',
        },
      },
    },
  },
  {
    name: 'teams_status',
    description: 'Check the current authentication status and session state.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'teams_send_message',
    description: 'Send a message to a Teams conversation. By default, sends to your own notes (self-chat). For channel threads, provide replyToMessageId to reply to an existing thread. Requires a valid session from prior login.',
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'The message content to send. Can include basic HTML formatting.',
        },
        conversationId: {
          type: 'string',
          description: 'The conversation ID to send to. Use "48:notes" for self-chat (default), or a channel/chat conversation ID.',
        },
        replyToMessageId: {
          type: 'string',
          description: 'For channel thread replies: the message ID of the thread root (the first message in the thread). When provided, the message is posted as a reply to that thread. Not needed for chats (1:1, group, meeting) as they are flat conversations.',
        },
      },
      required: ['content'],
    },
  },
  {
    name: 'teams_reply_to_thread',
    description: 'Reply to a channel message as a threaded reply. Use the conversationId and messageId from search results - the reply will appear under that message.',
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'The reply content to send.',
        },
        conversationId: {
          type: 'string',
          description: 'The channel conversation ID (from search results).',
        },
        messageId: {
          type: 'string',
          description: 'The message ID to reply to (from search results). This is the timestamp-based ID Teams uses for threading.',
        },
      },
      required: ['content', 'conversationId', 'messageId'],
    },
  },
  {
    name: 'teams_get_me',
    description: 'Get the current user\'s profile information including email, display name, and Teams ID. Useful for finding @mentions or identifying the current user.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'teams_search_people',
    description: 'Search for people in Microsoft Teams by name or email. Returns matching users with display name, email, job title, and department. Useful for finding someone to message.',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search term - can be a name, email address, or partial match',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of results to return (default: 10)',
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'teams_get_favorites',
    description: 'Get the user\'s favourite/pinned conversations in Teams. Returns conversation IDs with display names (channel name, chat topic, or participant names) and type (Channel, Chat, Meeting).',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'teams_add_favorite',
    description: 'Add a conversation to the user\'s favourites/pinned list.',
    inputSchema: {
      type: 'object',
      properties: {
        conversationId: {
          type: 'string',
          description: 'The conversation ID to pin (e.g., "19:abc@thread.tacv2")',
        },
      },
      required: ['conversationId'],
    },
  },
  {
    name: 'teams_remove_favorite',
    description: 'Remove a conversation from the user\'s favourites/pinned list.',
    inputSchema: {
      type: 'object',
      properties: {
        conversationId: {
          type: 'string',
          description: 'The conversation ID to unpin',
        },
      },
      required: ['conversationId'],
    },
  },
  {
    name: 'teams_save_message',
    description: 'Save (bookmark) a message in Teams. Saved messages can be accessed later from the Saved view in Teams.',
    inputSchema: {
      type: 'object',
      properties: {
        conversationId: {
          type: 'string',
          description: 'The conversation ID containing the message',
        },
        messageId: {
          type: 'string',
          description: 'The message ID to save (numeric string from search results)',
        },
      },
      required: ['conversationId', 'messageId'],
    },
  },
  {
    name: 'teams_unsave_message',
    description: 'Remove a saved (bookmarked) message in Teams.',
    inputSchema: {
      type: 'object',
      properties: {
        conversationId: {
          type: 'string',
          description: 'The conversation ID containing the message',
        },
        messageId: {
          type: 'string',
          description: 'The message ID to unsave',
        },
      },
      required: ['conversationId', 'messageId'],
    },
  },
  {
    name: 'teams_get_frequent_contacts',
    description: 'Get the user\'s frequently contacted people, ranked by interaction frequency. Useful for resolving ambiguous names (e.g., "Rob" → which Rob?) by checking who the user commonly works with. Returns display name, email, job title, and department.',
    inputSchema: {
      type: 'object',
      properties: {
        limit: {
          type: 'number',
          description: 'Maximum number of contacts to return (default: 50)',
        },
      },
    },
  },
  {
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
  },
  {
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
  },
  {
    name: 'teams_get_chat',
    description: 'Get the conversation ID for a 1:1 chat with a person. Use this to start a new chat or find an existing one. The conversation ID can then be used with teams_send_message to send messages.',
    inputSchema: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'The user\'s identifier. Can be: MRI (8:orgid:guid), object ID with tenant (guid@tenantId), or raw object ID (guid). Get this from teams_search_people results.',
        },
      },
      required: ['userId'],
    },
  },
];

// Input schemas for validation
const SearchInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  maxResults: z.number().optional().default(25),
  from: z.number().min(0).optional().default(0),
  size: z.number().min(1).max(100).optional().default(25),
});

const LoginInputSchema = z.object({
  forceNew: z.boolean().optional().default(false),
});

const SendMessageInputSchema = z.object({
  content: z.string().min(1, 'Message content cannot be empty'),
  conversationId: z.string().optional().default('48:notes'),
  replyToMessageId: z.string().optional(),
});

const ReplyToThreadInputSchema = z.object({
  content: z.string().min(1, 'Reply content cannot be empty'),
  conversationId: z.string().min(1, 'Conversation ID is required'),
  messageId: z.string().min(1, 'Message ID is required'),
});

const SearchPeopleInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  limit: z.number().min(1).max(50).optional().default(10),
});

const FavoriteInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

const SaveMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

const FrequentContactsInputSchema = z.object({
  limit: z.number().min(1).max(500).optional().default(50),
});

const GetThreadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  limit: z.number().min(1).max(200).optional().default(50),
});

const FindChannelInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  limit: z.number().min(1).max(50).optional().default(10),
});

const GetChatInputSchema = z.object({
  userId: z.string().min(1, 'User ID cannot be empty'),
});

/**
 * MCP Server for Teams integration.
 * 
 * Encapsulates all server state to allow multiple instances.
 */
export class TeamsServer {
  private browserManager: BrowserManager | null = null;
  private isInitialised = false;

  /**
   * Returns a standard MCP error response.
   */
  private formatError(error: McpError) {
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: error.message,
            errorCode: error.code,
            retryable: error.retryable,
            retryAfterMs: error.retryAfterMs,
            suggestions: error.suggestions,
          }, null, 2),
        },
      ],
      isError: true,
    };
  }

  /**
   * Returns a standard MCP success response.
   */
  private formatSuccess(data: Record<string, unknown>) {
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify({ success: true, ...data }, null, 2),
        },
      ],
    };
  }

  /**
   * Ensures the browser is running and authenticated.
   */
  private async ensureBrowser(headless: boolean = true): Promise<BrowserManager> {
    if (this.browserManager && this.isInitialised) {
      return this.browserManager;
    }

    // Close existing browser if any
    if (this.browserManager) {
      try {
        await closeBrowser(this.browserManager, true);
      } catch {
        // Ignore cleanup errors
      }
    }

    this.browserManager = await createBrowserContext({ headless });

    await ensureAuthenticated(
      this.browserManager.page,
      this.browserManager.context,
      (msg) => console.error(`[auth] ${msg}`)
    );

    this.isInitialised = true;
    return this.browserManager;
  }

  /**
   * Cleans up browser resources.
   */
  async cleanup(): Promise<void> {
    if (this.browserManager) {
      await closeBrowser(this.browserManager, true);
      this.browserManager = null;
      this.isInitialised = false;
    }
  }

  /**
   * Creates and configures the MCP server.
   */
  async createServer(): Promise<Server> {
    const server = new Server(
      {
        name: 'teams-mcp',
        version: '0.2.0',
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    // Handle resource listing
    server.setRequestHandler(ListResourcesRequestSchema, async () => {
      return {
        resources: [
          {
            uri: 'teams://me/profile',
            name: 'Current User Profile',
            description: 'The authenticated user\'s Teams profile including email and display name',
            mimeType: 'application/json',
          },
          {
            uri: 'teams://me/favorites',
            name: 'Pinned Conversations',
            description: 'The user\'s favourite/pinned Teams conversations',
            mimeType: 'application/json',
          },
          {
            uri: 'teams://status',
            name: 'Authentication Status',
            description: 'Current authentication status for all Teams APIs',
            mimeType: 'application/json',
          },
        ],
      };
    });

    // Handle resource reading
    server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const { uri } = request.params;

      switch (uri) {
        case 'teams://me/profile': {
          const profile = getUserProfile();
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(profile ?? { error: 'No valid session' }, null, 2),
              },
            ],
          };
        }

        case 'teams://me/favorites': {
          const result = await getFavorites();
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(
                  result.ok ? result.value.favorites : { error: result.error.message },
                  null,
                  2
                ),
              },
            ],
          };
        }

        case 'teams://status': {
          const tokenStatus = getSubstrateTokenStatus();
          const messageAuth = extractMessageAuth();
          const csaToken = extractCsaToken();

          const status = {
            directApi: {
              available: tokenStatus.hasToken,
              expiresAt: tokenStatus.expiresAt,
              minutesRemaining: tokenStatus.minutesRemaining,
            },
            messaging: {
              available: messageAuth !== null,
            },
            favorites: {
              available: messageAuth !== null && csaToken !== null,
            },
            session: {
              exists: hasSessionState(),
              likelyExpired: isSessionLikelyExpired(),
            },
          };

          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(status, null, 2),
              },
            ],
          };
        }

        default:
          throw new Error(`Unknown resource: ${uri}`);
      }
    });

    // Handle tool listing
    server.setRequestHandler(ListToolsRequestSchema, async () => {
      return { tools: TOOLS };
    });

    // Handle tool calls
    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'teams_search': {
            const input = SearchInputSchema.parse(args);

            // Try direct API first
            if (hasValidSubstrateToken()) {
              const result = await searchMessages(input.query, {
                maxResults: input.maxResults,
                from: input.from,
                size: input.size,
              });

              if (result.ok) {
                return this.formatSuccess({
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
                });
              }

              // Log error but fall through to browser
              console.error('[search] Direct API failed:', result.error.message);
            }

            // Fall back to browser-based search
            const manager = await this.ensureBrowser(false);

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
            await manager.page.waitForTimeout(3000);

            // Close browser after search
            await closeBrowser(manager, true);
            this.browserManager = null;
            this.isInitialised = false;

            return this.formatSuccess({
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
            });
          }

          case 'teams_login': {
            const input = LoginInputSchema.parse(args);

            if (this.browserManager) {
              await closeBrowser(this.browserManager, !input.forceNew);
              this.browserManager = null;
              this.isInitialised = false;
            }

            if (input.forceNew) {
              clearSessionState();
              clearTokenCache();
            }

            this.browserManager = await createBrowserContext({ headless: false });

            if (input.forceNew) {
              await forceNewLogin(
                this.browserManager.page,
                this.browserManager.context,
                (msg) => console.error(`[login] ${msg}`)
              );
            } else {
              await ensureAuthenticated(
                this.browserManager.page,
                this.browserManager.context,
                (msg) => console.error(`[login] ${msg}`)
              );
            }

            this.isInitialised = true;

            return this.formatSuccess({
              message: 'Login completed successfully. Session has been saved.',
            });
          }

          case 'teams_status': {
            const sessionExists = hasSessionState();
            const sessionExpired = isSessionLikelyExpired();
            const tokenStatus = getSubstrateTokenStatus();
            const messageAuth = extractMessageAuth();
            const csaToken = extractCsaToken();

            let authStatus = null;
            if (this.browserManager && this.isInitialised) {
              authStatus = await getAuthStatus(this.browserManager.page);
            }

            return this.formatSuccess({
              directApi: {
                available: tokenStatus.hasToken,
                expiresAt: tokenStatus.expiresAt,
                minutesRemaining: tokenStatus.minutesRemaining,
              },
              messaging: {
                available: messageAuth !== null,
              },
              favorites: {
                available: messageAuth !== null && csaToken !== null,
              },
              session: {
                exists: sessionExists,
                likelyExpired: sessionExpired,
              },
              browser: {
                running: this.browserManager !== null,
                initialised: this.isInitialised,
              },
              authentication: authStatus,
            });
          }

          case 'teams_send_message': {
            const input = SendMessageInputSchema.parse(args);

            const result = input.conversationId === '48:notes'
              ? await sendNoteToSelf(input.content)
              : await sendMessage(input.conversationId, input.content, {
                  replyToMessageId: input.replyToMessageId,
                });

            if (!result.ok) {
              return this.formatError(result.error);
            }

            // The timestamp is what Teams uses for threading - convert to string for use as threadReplyId
            const threadReplyId = result.value.timestamp ? String(result.value.timestamp) : undefined;

            const response: Record<string, unknown> = {
              messageId: result.value.messageId,
              timestamp: result.value.timestamp,
              conversationId: input.conversationId,
            };

            // Add threadReplyId for channel messages (needed to reply to this message later)
            if (threadReplyId && input.conversationId.includes('@thread.tacv2')) {
              response.threadReplyId = threadReplyId;
              response.note = 'Use threadReplyId (not messageId) if you want to reply to this message later.';
            }

            // Include replyToMessageId in response if this was a thread reply
            if (input.replyToMessageId) {
              response.replyToMessageId = input.replyToMessageId;
              response.note = 'Message posted as a reply to the thread.';
            }

            return this.formatSuccess(response);
          }

          case 'teams_reply_to_thread': {
            const input = ReplyToThreadInputSchema.parse(args);

            const result = await replyToThread(
              input.conversationId,
              input.messageId,
              input.content
            );

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              messageId: result.value.messageId,
              timestamp: result.value.timestamp,
              conversationId: result.value.conversationId,
              threadRootMessageId: result.value.threadRootMessageId,
              note: 'Reply posted to thread successfully.',
            });
          }

          case 'teams_get_me': {
            const profile = getUserProfile();

            if (!profile) {
              return this.formatError(createError(
                ErrorCode.AUTH_REQUIRED,
                'No valid session. Please use teams_login first.'
              ));
            }

            return this.formatSuccess({ profile });
          }

          case 'teams_search_people': {
            const input = SearchPeopleInputSchema.parse(args);

            const result = await searchPeople(input.query, input.limit);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              query: input.query,
              returned: result.value.returned,
              results: result.value.results,
            });
          }

          case 'teams_get_favorites': {
            const result = await getFavorites();

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              count: result.value.favorites.length,
              favorites: result.value.favorites,
            });
          }

          case 'teams_add_favorite': {
            const input = FavoriteInputSchema.parse(args);
            const result = await addFavorite(input.conversationId);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              message: `Added ${input.conversationId} to favourites`,
            });
          }

          case 'teams_remove_favorite': {
            const input = FavoriteInputSchema.parse(args);
            const result = await removeFavorite(input.conversationId);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              message: `Removed ${input.conversationId} from favourites`,
            });
          }

          case 'teams_save_message': {
            const input = SaveMessageInputSchema.parse(args);
            const result = await saveMessage(input.conversationId, input.messageId);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              message: 'Message saved',
              conversationId: input.conversationId,
              messageId: input.messageId,
            });
          }

          case 'teams_unsave_message': {
            const input = SaveMessageInputSchema.parse(args);
            const result = await unsaveMessage(input.conversationId, input.messageId);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              message: 'Message unsaved',
              conversationId: input.conversationId,
              messageId: input.messageId,
            });
          }

          case 'teams_get_frequent_contacts': {
            const input = FrequentContactsInputSchema.parse(args);

            const result = await getFrequentContacts(input.limit);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              returned: result.value.returned,
              contacts: result.value.results,
            });
          }

          case 'teams_get_thread': {
            const input = GetThreadInputSchema.parse(args);

            const result = await getThreadMessages(input.conversationId, { limit: input.limit });

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              conversationId: result.value.conversationId,
              messageCount: result.value.messages.length,
              messages: result.value.messages,
            });
          }

          case 'teams_find_channel': {
            const input = FindChannelInputSchema.parse(args);

            const result = await searchChannels(input.query, input.limit);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              query: input.query,
              count: result.value.returned,
              channels: result.value.results,
            });
          }

          case 'teams_get_chat': {
            const input = GetChatInputSchema.parse(args);

            const result = getOneOnOneChatId(input.userId);

            if (!result.ok) {
              return this.formatError(result.error);
            }

            return this.formatSuccess({
              conversationId: result.value.conversationId,
              otherUserId: result.value.otherUserId,
              currentUserId: result.value.currentUserId,
              note: 'Use this conversationId with teams_send_message to send a message. The conversation is created automatically when the first message is sent.',
            });
          }

          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);

        return this.formatError(createError(
          ErrorCode.UNKNOWN,
          message,
          { retryable: false }
        ));
      }
    });

    // Cleanup on server close
    server.onclose = async () => {
      await this.cleanup();
    };

    return server;
  }
}

/**
 * Creates and runs the MCP server.
 * Exported for backward compatibility.
 */
export async function createServer(): Promise<Server> {
  const teamsServer = new TeamsServer();
  return teamsServer.createServer();
}

/**
 * Runs the server with stdio transport.
 */
export async function runServer(): Promise<void> {
  const teamsServer = new TeamsServer();
  const server = await teamsServer.createServer();
  const transport = new StdioServerTransport();

  await server.connect(transport);

  // Handle shutdown signals
  const shutdown = async () => {
    await teamsServer.cleanup();
    process.exit(0);
  };

  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}
