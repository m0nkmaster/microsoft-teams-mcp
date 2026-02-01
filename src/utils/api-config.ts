/**
 * API endpoint configuration and header utilities.
 * 
 * Centralises all API URLs and common request headers.
 * 
 * Region is extracted from the user's session via DISCOVER-REGION-GTM,
 * so we no longer need validation - the data comes from Teams itself.
 */

import { NOTIFICATIONS_ID } from '../constants.js';

/** Substrate API endpoints. */
export const SUBSTRATE_API = {
  /** Full-text message search. */
  search: 'https://substrate.office.com/searchservice/api/v2/query',
  
  /** People and message typeahead suggestions. */
  suggestions: 'https://substrate.office.com/search/api/v1/suggestions',
  
  /** Frequent contacts list. */
  frequentContacts: 'https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache',
  
  /** People search. */
  peopleSearch: 'https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar',
  
  /** Channel search (org-wide, not just user's teams). */
  channelSearch: 'https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel',
} as const;

/** Chat service API endpoints. */
export const CHATSVC_API = {
  /**
   * Get messages URL for a conversation.
   * 
   * For thread replies in channels, provide replyToMessageId to append
   * `;messageid={id}` to the conversation path. This tells Teams the message
   * is a reply to an existing thread rather than a new top-level post.
   */
  messages: (region: string, conversationId: string, replyToMessageId?: string) => {
    // When replying to a thread, the URL includes ;messageid={threadRootId}
    const conversationPath = replyToMessageId
      ? `${conversationId};messageid=${replyToMessageId}`
      : conversationId;
    return `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationPath)}/messages`;
  },
  
  /** Get conversation metadata URL. */
  conversation: (region: string, conversationId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}`,
  
  /** Save/unsave message metadata URL. */
  messageMetadata: (region: string, conversationId: string, messageId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/rcmetadata/${messageId}`,
  
  /** Edit a specific message URL. */
  editMessage: (region: string, conversationId: string, messageId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}`,
  
  /** Delete a specific message URL (soft delete). */
  deleteMessage: (region: string, conversationId: string, messageId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}?behavior=softDelete`,
  
  /** Get consumption horizons (read receipts) for a thread. */
  consumptionHorizons: (region: string, threadId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/threads/${encodeURIComponent(threadId)}/consumptionhorizons`,
  
  /** Update consumption horizon (mark as read) for a conversation. */
  updateConsumptionHorizon: (region: string, conversationId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/properties?name=consumptionhorizon`,
  
  /** Activity feed (notifications) messages. */
  activityFeed: (region: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(NOTIFICATIONS_ID)}/messages`,
  
  /** Message emotions (reactions) URL. */
  messageEmotions: (region: string, conversationId: string, messageId: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}/properties?name=emotions`,
  
  /** Create a new thread (group chat). */
  createThread: (region: string) =>
    `https://teams.microsoft.com/api/chatsvc/${region}/v1/threads`,
} as const;

/** 
 * Calendar/Meeting API endpoints.
 * 
 * The mt/part endpoints use partitioned regions (e.g., amer-02, emea-03).
 * Some tenants use non-partitioned URLs (e.g., /api/mt/emea).
 * The correct format is extracted from the user's session via DISCOVER-REGION-GTM.
 */
export const CALENDAR_API = {
  /**
   * Get calendar view (meetings) for a date range.
   * 
   * Uses OData-style query parameters for filtering and pagination.
   * 
   * @param regionPartition - The full region-partition (e.g., "amer-02") or just region (e.g., "emea")
   * @param hasPartition - Whether the tenant uses partitioned URLs
   */
  calendarView: (regionPartition: string, hasPartition: boolean) =>
    hasPartition
      ? `https://teams.microsoft.com/api/mt/part/${regionPartition}/v2.1/me/calendars/calendarView`
      : `https://teams.microsoft.com/api/mt/${regionPartition}/v2.1/me/calendars/calendarView`,
} as const;

/** CSA (Chat Service Aggregator) API endpoints. */
export const CSA_API = {
  /** Conversation folders (favourites) URL. */
  conversationFolders: (region: string) =>
    `https://teams.microsoft.com/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`,
  
  /** Teams list (all teams/channels user is a member of). */
  teamsList: (region: string) =>
    `https://teams.microsoft.com/api/csa/${region}/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true`,
  
  /** Custom emoji metadata. */
  customEmojis: (region: string) =>
    `https://teams.microsoft.com/api/csa/${region}/api/v1/customemoji/metadata`,
} as const;

/** Common request headers for Teams API calls. */
export function getTeamsHeaders(): HeadersInit {
  return {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Origin': 'https://teams.microsoft.com',
    'Referer': 'https://teams.microsoft.com/',
  };
}

/** Headers for Bearer token authentication. */
export function getBearerHeaders(token: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authorization': `Bearer ${token}`,
  };
}

/** Headers for Skype token + Bearer authentication. */
export function getSkypeAuthHeaders(skypeToken: string, authToken: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${authToken}`,
  };
}

/** Headers for CSA API (Skype token + CSA Bearer). */
export function getCsaHeaders(skypeToken: string, csaToken: string): HeadersInit {
  return {
    ...getTeamsHeaders(),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${csaToken}`,
  };
}

/** Client version header for messaging API. */
export function getMessagingHeaders(skypeToken: string, authToken: string): HeadersInit {
  return {
    ...getSkypeAuthHeaders(skypeToken, authToken),
    'X-Ms-Client-Version': '1415/1.0.0.2025010401',
  };
}
