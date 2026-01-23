/**
 * MCP Server implementation for Teams search.
 * Exposes tools for searching Teams messages via browser automation.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

import { 
  createBrowserContext, 
  closeBrowser, 
  type BrowserManager 
} from './browser/context.js';
import { 
  ensureAuthenticated, 
  getAuthStatus, 
  forceNewLogin 
} from './browser/auth.js';
import { searchTeamsWithPagination } from './teams/search.js';
import { hasSessionState, isSessionLikelyExpired, clearSessionState } from './browser/session.js';
import { 
  directSearch, 
  hasValidToken, 
  getTokenStatus, 
  clearTokenCache,
  sendMessage,
  sendNoteToSelf,
  extractMessageAuth,
  extractCsaToken,
  getMe,
  searchPeople,
  getFrequentContacts,
  getFavorites,
  addFavorite,
  removeFavorite,
  saveMessage,
  unsaveMessage,
  getThreadMessages,
} from './teams/direct-api.js';

/** Returns a standard MCP error response for authentication failures. */
function authRequiredError() {
  return {
    content: [
      {
        type: 'text' as const,
        text: JSON.stringify({
          success: false,
          error: 'No valid authentication. Please use teams_login first.',
        }, null, 2),
      },
    ],
    isError: true,
  };
}

/** Returns a standard MCP error response for API failures. */
function apiError(error: string | undefined) {
  return {
    content: [
      {
        type: 'text' as const,
        text: JSON.stringify({
          success: false,
          error,
        }, null, 2),
      },
    ],
    isError: true,
  };
}

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
    description: 'Send a message to a Teams conversation. By default, sends to your own notes (self-chat). Requires a valid session from prior login.',
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'The message content to send. Can include basic HTML formatting.',
        },
        conversationId: {
          type: 'string',
          description: 'The conversation ID to send to. Use "48:notes" for self-chat (default), or a specific conversation ID.',
        },
      },
      required: ['content'],
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

// Server state
let browserManager: BrowserManager | null = null;
let isInitialised = false;

/**
 * Ensures the browser is running and authenticated.
 */
async function ensureBrowser(headless: boolean = true): Promise<BrowserManager> {
  if (browserManager && isInitialised) {
    return browserManager;
  }

  // Close existing browser if any
  if (browserManager) {
    try {
      await closeBrowser(browserManager, true);
    } catch {
      // Ignore cleanup errors
    }
  }

  browserManager = await createBrowserContext({ headless });
  
  await ensureAuthenticated(
    browserManager.page,
    browserManager.context,
    (msg) => console.error(`[auth] ${msg}`)
  );

  isInitialised = true;
  return browserManager;
}

/**
 * Creates and runs the MCP server.
 */
export async function createServer(): Promise<Server> {
  const server = new Server(
    {
      name: 'teams-mcp',
      version: '0.1.0',
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

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
          
          // Try direct API first (no browser needed)
          if (hasValidToken()) {
            try {
              const { results, pagination } = await directSearch(
                input.query,
                { 
                  maxResults: input.maxResults,
                  from: input.from,
                  size: input.size,
                }
              );

              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      success: true,
                      mode: 'direct-api',
                      query: input.query,
                      resultCount: results.length,
                      pagination: {
                        from: pagination.from,
                        size: pagination.size,
                        returned: pagination.returned,
                        total: pagination.total,
                        hasMore: pagination.hasMore,
                        nextFrom: pagination.hasMore ? pagination.from + pagination.returned : undefined,
                      },
                      results,
                    }, null, 2),
                  },
                ],
              };
            } catch (error) {
              // Token might be expired, fall through to browser-based login
              console.error('[search] Direct API failed:', error instanceof Error ? error.message : error);
            }
          }
          
          // No valid token - need browser login
          // Open visible browser for user to log in
          const manager = await ensureBrowser(false); // visible browser
          
          const { results, pagination } = await searchTeamsWithPagination(
            manager.page,
            input.query,
            { 
              maxResults: input.maxResults,
              from: input.from,
              size: input.size,
            }
          );

          // Wait for MSAL to store the search token after the API call
          await manager.page.waitForTimeout(3000);

          // After successful browser search, close the browser
          // The session state is saved, so next time we can use direct API
          await closeBrowser(manager, true);
          browserManager = null;
          isInitialised = false;

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
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
                    nextFrom: pagination.hasMore ? pagination.from + pagination.returned : undefined,
                  },
                  results,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_login': {
          const input = LoginInputSchema.parse(args);
          
          // For login, we need a visible browser
          if (browserManager) {
            await closeBrowser(browserManager, !input.forceNew);
            browserManager = null;
            isInitialised = false;
          }

          if (input.forceNew) {
            clearSessionState();
            clearTokenCache();
          }

          // Create visible browser for login
          browserManager = await createBrowserContext({ headless: false });
          
          if (input.forceNew) {
            await forceNewLogin(
              browserManager.page,
              browserManager.context,
              (msg) => console.error(`[login] ${msg}`)
            );
          } else {
            await ensureAuthenticated(
              browserManager.page,
              browserManager.context,
              (msg) => console.error(`[login] ${msg}`)
            );
          }

          isInitialised = true;

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: 'Login completed successfully. Session has been saved.',
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_status': {
          const sessionExists = hasSessionState();
          const sessionExpired = isSessionLikelyExpired();
          const tokenStatus = getTokenStatus();
          const messageAuth = extractMessageAuth();
          const csaToken = extractCsaToken();
          
          let authStatus = null;
          if (browserManager && isInitialised) {
            authStatus = await getAuthStatus(browserManager.page);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
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
                    running: browserManager !== null,
                    initialised: isInitialised,
                  },
                  authentication: authStatus,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_send_message': {
          const input = SendMessageInputSchema.parse(args);
          
          // Check if we have valid message auth
          if (!extractMessageAuth()) {
            return authRequiredError();
          }

          const result = input.conversationId === '48:notes'
            ? await sendNoteToSelf(input.content)
            : await sendMessage(input.conversationId, input.content);

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  messageId: result.messageId,
                  timestamp: result.timestamp,
                  conversationId: input.conversationId,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_get_me': {
          const profile = getMe();
          
          if (!profile) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    success: false,
                    error: 'No valid session. Please use teams_login first.',
                  }, null, 2),
                },
              ],
              isError: true,
            };
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  profile,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_search_people': {
          const input = SearchPeopleInputSchema.parse(args);
          
          // Check if we have a valid token
          if (!hasValidToken()) {
            return authRequiredError();
          }

          const { results, returned } = await searchPeople(input.query, input.limit);

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  query: input.query,
                  returned,
                  results,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_get_favorites': {
          const result = await getFavorites();

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  count: result.favorites.length,
                  favorites: result.favorites,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_add_favorite': {
          const input = FavoriteInputSchema.parse(args);
          const result = await addFavorite(input.conversationId);

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `Added ${input.conversationId} to favourites`,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_remove_favorite': {
          const input = FavoriteInputSchema.parse(args);
          const result = await removeFavorite(input.conversationId);

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `Removed ${input.conversationId} from favourites`,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_save_message': {
          const input = SaveMessageInputSchema.parse(args);
          const result = await saveMessage(input.conversationId, input.messageId);

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: 'Message saved',
                  conversationId: input.conversationId,
                  messageId: input.messageId,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_unsave_message': {
          const input = SaveMessageInputSchema.parse(args);
          const result = await unsaveMessage(input.conversationId, input.messageId);

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: 'Message unsaved',
                  conversationId: input.conversationId,
                  messageId: input.messageId,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_get_frequent_contacts': {
          const input = FrequentContactsInputSchema.parse(args);
          
          // Check if we have a valid token
          if (!hasValidToken()) {
            return authRequiredError();
          }

          const { contacts, returned } = await getFrequentContacts(input.limit);

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  returned,
                  contacts,
                }, null, 2),
              },
            ],
          };
        }

        case 'teams_get_thread': {
          const input = GetThreadInputSchema.parse(args);
          
          const result = await getThreadMessages(input.conversationId, { limit: input.limit });

          if (!result.success) {
            return apiError(result.error);
          }

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  conversationId: result.conversationId,
                  messageCount: result.messages?.length ?? 0,
                  messages: result.messages,
                }, null, 2),
              },
            ],
          };
        }

        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: false,
              error: message,
            }, null, 2),
          },
        ],
        isError: true,
      };
    }
  });

  // Cleanup on server close
  server.onclose = async () => {
    if (browserManager) {
      await closeBrowser(browserManager, true);
      browserManager = null;
      isInitialised = false;
    }
  };

  return server;
}

/**
 * Runs the server with stdio transport.
 */
export async function runServer(): Promise<void> {
  const server = await createServer();
  const transport = new StdioServerTransport();
  
  await server.connect(transport);
  
  // Handle shutdown signals
  const shutdown = async () => {
    if (browserManager) {
      await closeBrowser(browserManager, true);
    }
    process.exit(0);
  };

  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}
