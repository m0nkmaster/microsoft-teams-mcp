/**
 * Chat Service API client for messaging operations.
 * 
 * Handles all calls to teams.microsoft.com/api/chatsvc endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getMessagingHeaders, getSkypeAuthHeaders, getTeamsHeaders, validateRegion, type Region } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractMessageAuth,
  getUserDisplayName,
} from '../auth/token-extractor.js';
import { stripHtml, buildMessageLink, buildOneOnOneConversationId, extractObjectId } from '../utils/parsers.js';

/** Result of sending a message. */
export interface SendMessageResult {
  messageId: string;
  timestamp?: number;
}

/** A message from a thread/conversation. */
export interface ThreadMessage {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  conversationId: string;
  clientMessageId?: string;
  isFromMe?: boolean;
  messageLink?: string;
}

/** Result of getting thread messages. */
export interface GetThreadResult {
  conversationId: string;
  messages: ThreadMessage[];
}

/** Result of saving/unsaving a message. */
export interface SaveMessageResult {
  conversationId: string;
  messageId: string;
  saved: boolean;
}

/** Options for sending a message. */
export interface SendMessageOptions {
  /** Region for the API call (default: 'amer'). */
  region?: string;
  /**
   * Message ID of the thread root to reply to.
   * 
   * When provided, the message is posted as a reply to an existing thread
   * in a channel. The conversationId should be the channel ID, and this
   * should be the ID of the first message in the thread.
   * 
   * For chats (1:1, group, meeting), this is not needed - all messages
   * are part of the same flat conversation.
   */
  replyToMessageId?: string;
}

/**
 * Sends a message to a Teams conversation.
 * 
 * For channels, you can either:
 * - Post a new top-level message: just provide the channel's conversationId
 * - Reply to a thread: provide the channel's conversationId AND replyToMessageId
 * 
 * For chats (1:1, group, meeting), all messages go to the same conversation
 * without threading - just provide the conversationId.
 */
export async function sendMessage(
  conversationId: string,
  content: string,
  options: SendMessageOptions = {}
): Promise<Result<SendMessageResult>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const { region = 'amer', replyToMessageId } = options;
  const validRegion = validateRegion(region);
  const displayName = getUserDisplayName() || 'User';

  // Generate unique message ID
  const clientMessageId = Date.now().toString();

  // Wrap content in paragraph if not already HTML
  const htmlContent = content.startsWith('<') ? content : `<p>${escapeHtml(content)}</p>`;

  const body = {
    content: htmlContent,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    imdisplayname: displayName,
    clientmessageid: clientMessageId,
  };

  const url = CHATSVC_API.messages(validRegion, conversationId, replyToMessageId);

  const response = await httpRequest<{ OriginalArrivalTime?: number }>(
    url,
    {
      method: 'POST',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId: clientMessageId,
    timestamp: response.value.data.OriginalArrivalTime,
  });
}

/**
 * Sends a message to your own notes/self-chat.
 */
export async function sendNoteToSelf(content: string): Promise<Result<SendMessageResult>> {
  return sendMessage('48:notes', content);
}

/** Result of replying to a thread. */
export interface ReplyToThreadResult extends SendMessageResult {
  /** The thread root message ID used for the reply. */
  threadRootMessageId: string;
  /** The conversation ID (channel) the reply was posted to. */
  conversationId: string;
}

/**
 * Replies to a thread in a Teams channel.
 * 
 * Uses the provided messageId directly as the thread root. In Teams channels:
 * - If messageId is a top-level post, the reply goes under that post
 * - If messageId is already a reply within a thread, the reply goes to the same thread
 * 
 * For channel messages from search results, the messageId is typically the thread root
 * (the original message that started the thread).
 * 
 * @param conversationId - The channel conversation ID (from search results)
 * @param messageId - The message ID to reply to (typically the thread root from search)
 * @param content - The reply content
 * @param region - API region (default: 'amer')
 * @returns The result including the new message ID
 */
export async function replyToThread(
  conversationId: string,
  messageId: string,
  content: string,
  region: string = 'amer'
): Promise<Result<ReplyToThreadResult>> {
  // Use the provided messageId directly as the thread root
  // Search results return the thread root ID for channel messages
  const threadRootMessageId = messageId;
  
  // Send the reply using the provided message ID as the thread root
  const sendResult = await sendMessage(conversationId, content, {
    region,
    replyToMessageId: threadRootMessageId,
  });
  
  if (!sendResult.ok) {
    return sendResult;
  }
  
  return ok({
    messageId: sendResult.value.messageId,
    timestamp: sendResult.value.timestamp,
    threadRootMessageId,
    conversationId,
  });
}

/**
 * Gets messages from a Teams conversation/thread.
 */
export async function getThreadMessages(
  conversationId: string,
  options: { limit?: number; startTime?: number } = {},
  region: string = 'amer'
): Promise<Result<GetThreadResult>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(validRegion, conversationId);
  url += `?view=msnp24Equivalent&pageSize=${limit}`;

  if (options.startTime) {
    url += `&startTime=${options.startTime}`;
  }

  const response = await httpRequest<{ messages?: unknown[] }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken),
    }
  );

  if (!response.ok) {
    return response;
  }

  const rawMessages = response.value.data.messages;
  if (!Array.isArray(rawMessages)) {
    return ok({
      conversationId,
      messages: [],
    });
  }

  const messages: ThreadMessage[] = [];

  for (const raw of rawMessages) {
    const msg = raw as Record<string, unknown>;

    // Skip non-message types
    const messageType = msg.messagetype as string;
    if (!messageType || messageType.startsWith('Control/') || messageType === 'ThreadActivity/AddMember') {
      continue;
    }

    const id = msg.id as string || msg.originalarrivaltime as string;
    if (!id) continue;

    const content = msg.content as string || '';
    const contentType = msg.messagetype as string || 'Text';

    const fromMri = msg.from as string || '';
    const displayName = msg.imdisplayname as string || msg.displayName as string;

    const timestamp = msg.originalarrivaltime as string ||
      msg.composetime as string ||
      new Date(parseInt(id, 10)).toISOString();

    // Build message link
    const messageLink = /^\d+$/.test(id)
      ? buildMessageLink(conversationId, id)
      : undefined;

    messages.push({
      id,
      content: stripHtml(content),
      contentType,
      sender: {
        mri: fromMri,
        displayName,
      },
      timestamp,
      conversationId,
      clientMessageId: msg.clientmessageid as string,
      isFromMe: fromMri === auth.userMri,
      messageLink,
    });
  }

  // Sort by timestamp (oldest first)
  messages.sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());

  return ok({
    conversationId,
    messages,
  });
}

/**
 * Saves (bookmarks) a message.
 */
export async function saveMessage(
  conversationId: string,
  messageId: string,
  region: string = 'amer'
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, true, region);
}

/**
 * Unsaves (removes bookmark from) a message.
 */
export async function unsaveMessage(
  conversationId: string,
  messageId: string,
  region: string = 'amer'
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, false, region);
}

/**
 * Internal function to set the saved state of a message.
 */
async function setMessageSavedState(
  conversationId: string,
  messageId: string,
  saved: boolean,
  region: string
): Promise<Result<SaveMessageResult>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const url = CHATSVC_API.messageMetadata(validRegion, conversationId, messageId);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken),
      body: JSON.stringify({
        s: saved ? 1 : 0,
        mid: parseInt(messageId, 10),
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    saved,
  });
}

/** Result of editing a message. */
export interface EditMessageResult {
  messageId: string;
  conversationId: string;
}

/** Result of deleting a message. */
export interface DeleteMessageResult {
  messageId: string;
  conversationId: string;
}

/**
 * Edits an existing message.
 * 
 * Note: You can only edit your own messages. The API will reject
 * attempts to edit messages from other users.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The ID of the message to edit
 * @param newContent - The new content for the message
 * @param region - API region (default: 'amer')
 */
export async function editMessage(
  conversationId: string,
  messageId: string,
  newContent: string,
  region: string = 'amer'
): Promise<Result<EditMessageResult>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const displayName = getUserDisplayName() || 'User';
  
  // Wrap content in paragraph if not already HTML
  const htmlContent = newContent.startsWith('<') ? newContent : `<p>${escapeHtml(newContent)}</p>`;

  // Build the edit request body
  // The API requires the message structure with updated content
  const body = {
    id: messageId,
    type: 'Message',
    conversationid: conversationId,
    content: htmlContent,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    imdisplayname: displayName,
  };

  const url = CHATSVC_API.editMessage(validRegion, conversationId, messageId);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId,
    conversationId,
  });
}

/**
 * Deletes a message (soft delete).
 * 
 * Note: You can only delete your own messages, unless you are a
 * channel owner/moderator. The API will reject unauthorised attempts.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The ID of the message to delete
 * @param region - API region (default: 'amer')
 */
export async function deleteMessage(
  conversationId: string,
  messageId: string,
  region: string = 'amer'
): Promise<Result<DeleteMessageResult>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const url = CHATSVC_API.deleteMessage(validRegion, conversationId, messageId);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'DELETE',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId,
    conversationId,
  });
}

/**
 * Gets properties for a single conversation.
 */
export async function getConversationProperties(
  conversationId: string,
  region: string = 'amer'
): Promise<Result<{ displayName?: string; conversationType?: string }>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const url = CHATSVC_API.conversation(validRegion, conversationId) + '?view=msnp24Equivalent';

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const threadProps = data.threadProperties as Record<string, unknown> | undefined;
  const productType = threadProps?.productThreadType as string | undefined;

  // Try to get display name from various sources
  let displayName: string | undefined;

  if (threadProps?.topicThreadTopic) {
    displayName = threadProps.topicThreadTopic as string;
  }

  if (!displayName && threadProps?.topic) {
    displayName = threadProps.topic as string;
  }

  if (!displayName && threadProps?.spaceThreadTopic) {
    displayName = threadProps.spaceThreadTopic as string;
  }

  if (!displayName && threadProps?.threadtopic) {
    displayName = threadProps.threadtopic as string;
  }

  // For chats without a topic: build from members
  if (!displayName) {
    const members = data.members as Array<Record<string, unknown>> | undefined;
    if (members && members.length > 0) {
      const otherMembers = members
        .filter(m => m.mri !== auth.userMri && m.id !== auth.userMri)
        .map(m => (m.friendlyName || m.displayName || m.name) as string | undefined)
        .filter((name): name is string => !!name);

      if (otherMembers.length > 0) {
        displayName = otherMembers.length <= 3
          ? otherMembers.join(', ')
          : `${otherMembers.slice(0, 3).join(', ')} + ${otherMembers.length - 3} more`;
      }
    }
  }

  // Determine conversation type
  let conversationType: string | undefined;

  if (productType) {
    if (productType === 'Meeting') {
      conversationType = 'Meeting';
    } else if (productType.includes('Channel') || productType === 'TeamsTeam') {
      conversationType = 'Channel';
    } else if (productType === 'Chat' || productType === 'OneOnOne') {
      conversationType = 'Chat';
    }
  }

  // Fallback to ID pattern detection
  if (!conversationType) {
    if (conversationId.includes('meeting_')) {
      conversationType = 'Meeting';
    } else if (threadProps?.groupId) {
      conversationType = 'Channel';
    } else if (conversationId.includes('@thread.tacv2') || conversationId.includes('@thread.v2')) {
      conversationType = 'Chat';
    } else if (conversationId.startsWith('8:')) {
      conversationType = 'Chat';
    }
  }

  return ok({ displayName, conversationType });
}

/**
 * Extracts unique participant names from recent messages.
 */
export async function extractParticipantNames(
  conversationId: string,
  region: string = 'amer'
): Promise<Result<string | undefined>> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  let url = CHATSVC_API.messages(validRegion, conversationId);
  url += '?view=msnp24Equivalent&pageSize=10';

  const response = await httpRequest<{ messages?: unknown[] }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken),
    }
  );

  if (!response.ok) {
    return ok(undefined);
  }

  const messages = response.value.data.messages;
  if (!messages || messages.length === 0) {
    return ok(undefined);
  }

  const senderNames = new Set<string>();
  for (const msg of messages) {
    const m = msg as Record<string, unknown>;
    const fromMri = m.from as string || '';
    const displayName = m.imdisplayname as string;

    if (fromMri === auth.userMri || !displayName) {
      continue;
    }

    senderNames.add(displayName);
  }

  if (senderNames.size === 0) {
    return ok(undefined);
  }

  const names = Array.from(senderNames);
  const result = names.length <= 3
    ? names.join(', ')
    : `${names.slice(0, 3).join(', ')} + ${names.length - 3} more`;

  return ok(result);
}

/**
 * Escapes HTML special characters.
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Result of getting a 1:1 conversation. */
export interface GetOneOnOneChatResult {
  conversationId: string;
  otherUserId: string;
  currentUserId: string;
}

/**
 * Gets the conversation ID for a 1:1 chat with another user.
 * 
 * This constructs the predictable conversation ID format used by Teams
 * for 1:1 chats. The conversation ID is: `19:{id1}_{id2}@unq.gbl.spaces`
 * where id1 and id2 are the two users' object IDs sorted lexicographically.
 * 
 * Note: This doesn't create the conversation - it just returns the ID.
 * The conversation is implicitly created when the first message is sent.
 * 
 * @param otherUserIdentifier - The other user's MRI, object ID, or ID with tenant
 * @returns The conversation ID, or an error if auth is missing or ID is invalid
 */
export function getOneOnOneChatId(
  otherUserIdentifier: string
): Result<GetOneOnOneChatResult> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  // Extract the current user's object ID from their MRI
  const currentUserId = extractObjectId(auth.userMri);
  if (!currentUserId) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not extract user ID from session. Please try logging in again.'
    ));
  }

  // Extract the other user's object ID
  const otherUserId = extractObjectId(otherUserIdentifier);
  if (!otherUserId) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `Invalid user identifier: ${otherUserIdentifier}. Expected MRI (8:orgid:guid), ID with tenant (guid@tenant), or raw GUID.`
    ));
  }

  const conversationId = buildOneOnOneConversationId(currentUserId, otherUserId);

  if (!conversationId) {
    // This shouldn't happen if both IDs were validated above, but handle it anyway
    return err(createError(
      ErrorCode.UNKNOWN,
      'Failed to construct conversation ID.'
    ));
  }

  return ok({
    conversationId,
    otherUserId,
    currentUserId,
  });
}
