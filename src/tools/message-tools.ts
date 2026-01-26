/**
 * Messaging-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import {
  sendMessage,
  sendNoteToSelf,
  replyToThread,
  saveMessage,
  unsaveMessage,
  getOneOnOneChatId,
  editMessage,
  deleteMessage,
  getUnreadStatus,
  markAsRead,
  getActivityFeed,
} from '../api/chatsvc-api.js';
import { getFavorites, addFavorite, removeFavorite } from '../api/csa-api.js';
import { SELF_CHAT_ID, MAX_UNREAD_AGGREGATE_CHECK } from '../constants.js';
import { ErrorCode } from '../types/errors.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SendMessageInputSchema = z.object({
  content: z.string().min(1, 'Message content cannot be empty'),
  conversationId: z.string().optional().default(SELF_CHAT_ID),
  replyToMessageId: z.string().optional(),
});

export const ReplyToThreadInputSchema = z.object({
  content: z.string().min(1, 'Reply content cannot be empty'),
  conversationId: z.string().min(1, 'Conversation ID is required'),
  messageId: z.string().min(1, 'Message ID is required'),
});

export const FavoriteInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const SaveMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetChatInputSchema = z.object({
  userId: z.string().min(1, 'User ID cannot be empty'),
});

export const EditMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  content: z.string().min(1, 'Content cannot be empty'),
});

export const DeleteMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetUnreadInputSchema = z.object({
  conversationId: z.string().optional(),
});

export const MarkAsReadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetActivityInputSchema = z.object({
  limit: z.number().min(1).max(200).optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const sendMessageToolDefinition: Tool = {
  name: 'teams_send_message',
  description: 'Send a message to a Teams conversation. By default, sends to your own notes (self-chat). For channel thread replies, use teams_reply_to_thread instead (simpler). For chats (1:1, group, meeting), just provide the conversationId.',
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
        description: 'For channel thread replies (advanced). Prefer using teams_reply_to_thread instead.',
      },
    },
    required: ['content'],
  },
};

const replyToThreadToolDefinition: Tool = {
  name: 'teams_reply_to_thread',
  description: 'Reply to a channel message as a threaded reply. Use the conversationId and messageId from search results, or conversationId and threadReplyId from a previous teams_send_message response.',
  inputSchema: {
    type: 'object',
    properties: {
      content: {
        type: 'string',
        description: 'The reply content to send.',
      },
      conversationId: {
        type: 'string',
        description: 'The channel conversation ID.',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to reply to. Use messageId from search results, or threadReplyId from a teams_send_message response.',
      },
    },
    required: ['content', 'conversationId', 'messageId'],
  },
};

const getFavoritesToolDefinition: Tool = {
  name: 'teams_get_favorites',
  description: 'Get the user\'s favourite/pinned conversations in Teams. Returns conversation IDs with display names (channel name, chat topic, or participant names) and type (Channel, Chat, Meeting).',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

const addFavoriteToolDefinition: Tool = {
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
};

const removeFavoriteToolDefinition: Tool = {
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
};

const saveMessageToolDefinition: Tool = {
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
};

const unsaveMessageToolDefinition: Tool = {
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
};

const getChatToolDefinition: Tool = {
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
};

const editMessageToolDefinition: Tool = {
  name: 'teams_edit_message',
  description: 'Edit one of your own messages. You can only edit messages you sent. The API will reject attempts to edit other users\' messages.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to edit (numeric string from search results or teams_get_thread)',
      },
      content: {
        type: 'string',
        description: 'The new content for the message. Can include basic HTML formatting.',
      },
    },
    required: ['conversationId', 'messageId', 'content'],
  },
};

const deleteMessageToolDefinition: Tool = {
  name: 'teams_delete_message',
  description: 'Delete one of your own messages (soft delete - the message remains but content becomes empty). You can only delete messages you sent, unless you are a channel owner/moderator.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to delete (numeric string from search results or teams_get_thread)',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getUnreadToolDefinition: Tool = {
  name: 'teams_get_unread',
  description: 'Get unread message status. Without parameters, returns aggregate unread counts across all favourite/pinned conversations. With a conversationId, returns unread status for that specific conversation.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'Optional. A specific conversation ID to check. If omitted, checks all favourites.',
      },
    },
  },
};

const markAsReadToolDefinition: Tool = {
  name: 'teams_mark_read',
  description: 'Mark a conversation as read up to a specific message. This updates your read position so messages up to (and including) the specified message are marked as read.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to mark as read',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to mark as read up to (all messages up to this point will be marked read)',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getActivityToolDefinition: Tool = {
  name: 'teams_get_activity',
  description: 'Get the user\'s activity feed - mentions, reactions, replies, and other notifications. Returns recent activity items with sender, content, and source conversation context.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of activity items to return (default: 50, max: 200)',
      },
    },
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleSendMessage(
  input: z.infer<typeof SendMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = input.conversationId === SELF_CHAT_ID
    ? await sendNoteToSelf(input.content)
    : await sendMessage(input.conversationId, input.content, {
        replyToMessageId: input.replyToMessageId,
      });

  if (!result.ok) {
    return { success: false, error: result.error };
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

  return { success: true, data: response };
}

async function handleReplyToThread(
  input: z.infer<typeof ReplyToThreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await replyToThread(
    input.conversationId,
    input.messageId,
    input.content
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      messageId: result.value.messageId,
      timestamp: result.value.timestamp,
      conversationId: result.value.conversationId,
      threadRootMessageId: result.value.threadRootMessageId,
      note: 'Reply posted to thread successfully.',
    },
  };
}

async function handleGetFavorites(
  _input: Record<string, never>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFavorites();

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.favorites.length,
      favorites: result.value.favorites,
    },
  };
}

async function handleAddFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Added ${input.conversationId} to favourites`,
    },
  };
}

async function handleRemoveFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Removed ${input.conversationId} from favourites`,
    },
  };
}

async function handleSaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await saveMessage(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message saved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleUnsaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await unsaveMessage(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message unsaved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleGetChat(
  input: z.infer<typeof GetChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = getOneOnOneChatId(input.userId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      otherUserId: result.value.otherUserId,
      currentUserId: result.value.currentUserId,
      note: 'Use this conversationId with teams_send_message to send a message. The conversation is created automatically when the first message is sent.',
    },
  };
}

async function handleEditMessage(
  input: z.infer<typeof EditMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await editMessage(
    input.conversationId,
    input.messageId,
    input.content
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message edited successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleDeleteMessage(
  input: z.infer<typeof DeleteMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await deleteMessage(
    input.conversationId,
    input.messageId
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message deleted successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleGetUnread(
  input: z.infer<typeof GetUnreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  // If a specific conversation is provided, just check that one
  if (input.conversationId) {
    const result = await getUnreadStatus(input.conversationId);
    if (!result.ok) {
      return { success: false, error: result.error };
    }

    return {
      success: true,
      data: {
        conversationId: result.value.conversationId,
        unreadCount: result.value.unreadCount,
        lastReadMessageId: result.value.lastReadMessageId,
        latestMessageId: result.value.latestMessageId,
      },
    };
  }

  // Aggregate mode: check all favourites
  const favResult = await getFavorites();
  if (!favResult.ok) {
    return { success: false, error: favResult.error };
  }

  const favorites = favResult.value.favorites;
  const conversations: Array<{
    conversationId: string;
    displayName?: string;
    conversationType?: string;
    unreadCount: number;
  }> = [];

  let totalUnread = 0;
  let checkedCount = 0;
  let errorCount = 0;

  // Check unread status for each favourite (limit to prevent timeout)
  const maxToCheck = MAX_UNREAD_AGGREGATE_CHECK;
  for (const fav of favorites.slice(0, maxToCheck)) {
    const unreadResult = await getUnreadStatus(fav.conversationId);
    checkedCount++;

    if (unreadResult.ok) {
      if (unreadResult.value.unreadCount > 0) {
        conversations.push({
          conversationId: fav.conversationId,
          displayName: fav.displayName,
          conversationType: fav.conversationType,
          unreadCount: unreadResult.value.unreadCount,
        });
        totalUnread += unreadResult.value.unreadCount;
      }
    } else {
      errorCount++;
    }
  }

  // If all checks failed, return an error rather than misleading success
  if (checkedCount > 0 && errorCount === checkedCount) {
    return {
      success: false,
      error: {
        code: ErrorCode.API_ERROR,
        message: `Failed to check unread status for all ${checkedCount} favourites`,
        retryable: true,
        suggestions: ['Check authentication status with teams_status', 'Try teams_login to refresh session'],
      },
    };
  }

  return {
    success: true,
    data: {
      totalUnread,
      conversationsWithUnread: conversations.length,
      conversations,
      checked: checkedCount,
      totalFavorites: favorites.length,
      errors: errorCount > 0 ? errorCount : undefined,
      note: favorites.length > maxToCheck
        ? `Checked first ${maxToCheck} of ${favorites.length} favourites`
        : undefined,
    },
  };
}

async function handleMarkAsRead(
  input: z.infer<typeof MarkAsReadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await markAsRead(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Conversation marked as read',
      conversationId: result.value.conversationId,
      markedUpTo: result.value.markedUpTo,
    },
  };
}

async function handleGetActivity(
  input: z.infer<typeof GetActivityInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getActivityFeed({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.activities.length,
      activities: result.value.activities,
      syncState: result.value.syncState,
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const sendMessageTool: RegisteredTool<typeof SendMessageInputSchema> = {
  definition: sendMessageToolDefinition,
  schema: SendMessageInputSchema,
  handler: handleSendMessage,
};

export const replyToThreadTool: RegisteredTool<typeof ReplyToThreadInputSchema> = {
  definition: replyToThreadToolDefinition,
  schema: ReplyToThreadInputSchema,
  handler: handleReplyToThread,
};

export const getFavoritesTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: getFavoritesToolDefinition,
  schema: z.object({}),
  handler: handleGetFavorites,
};

export const addFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: addFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleAddFavorite,
};

export const removeFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: removeFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleRemoveFavorite,
};

export const saveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: saveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleSaveMessage,
};

export const unsaveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: unsaveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleUnsaveMessage,
};

export const getChatTool: RegisteredTool<typeof GetChatInputSchema> = {
  definition: getChatToolDefinition,
  schema: GetChatInputSchema,
  handler: handleGetChat,
};

export const editMessageTool: RegisteredTool<typeof EditMessageInputSchema> = {
  definition: editMessageToolDefinition,
  schema: EditMessageInputSchema,
  handler: handleEditMessage,
};

export const deleteMessageTool: RegisteredTool<typeof DeleteMessageInputSchema> = {
  definition: deleteMessageToolDefinition,
  schema: DeleteMessageInputSchema,
  handler: handleDeleteMessage,
};

export const getUnreadTool: RegisteredTool<typeof GetUnreadInputSchema> = {
  definition: getUnreadToolDefinition,
  schema: GetUnreadInputSchema,
  handler: handleGetUnread,
};

export const markAsReadTool: RegisteredTool<typeof MarkAsReadInputSchema> = {
  definition: markAsReadToolDefinition,
  schema: MarkAsReadInputSchema,
  handler: handleMarkAsRead,
};

export const getActivityTool: RegisteredTool<typeof GetActivityInputSchema> = {
  definition: getActivityToolDefinition,
  schema: GetActivityInputSchema,
  handler: handleGetActivity,
};

/** All message-related tools. */
export const messageTools = [
  sendMessageTool,
  replyToThreadTool,
  getFavoritesTool,
  addFavoriteTool,
  removeFavoriteTool,
  saveMessageTool,
  unsaveMessageTool,
  getChatTool,
  editMessageTool,
  deleteMessageTool,
  getUnreadTool,
  markAsReadTool,
  getActivityTool,
];
