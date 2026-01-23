/**
 * Pure parsing functions for Teams API responses.
 * 
 * These functions transform raw API responses into our internal types.
 * They are extracted here for testability - no side effects or external dependencies.
 */

import type { TeamsSearchResult } from '../types/teams.js';

/** Person search result from Substrate suggestions API. */
export interface PersonSearchResult {
  id: string;              // Azure AD object ID
  mri: string;             // Teams MRI (8:orgid:guid)
  displayName: string;     // Full display name
  email?: string;          // Primary email address
  givenName?: string;      // First name
  surname?: string;        // Last name
  jobTitle?: string;       // Job title
  department?: string;     // Department
  companyName?: string;    // Company name
}

/** User profile extracted from JWT tokens. */
export interface UserProfile {
  id: string;           // Azure AD object ID (oid)
  mri: string;          // Teams MRI (8:orgid:guid)
  email: string;        // User principal name / email
  displayName: string;  // Full display name
  givenName?: string;   // First name
  surname?: string;     // Last name
  tenantId?: string;    // Azure tenant ID
}

/**
 * Strips HTML tags from content for display.
 */
export function stripHtml(html: string): string {
  return html
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Builds a deep link to open a message in Teams.
 * 
 * Format: https://teams.microsoft.com/l/message/{conversationId}/{messageTimestamp}
 * 
 * @param conversationId - The conversation/thread ID (e.g., "19:xxx@thread.tacv2")
 * @param messageTimestamp - The message timestamp in epoch milliseconds
 */
export function buildMessageLink(
  conversationId: string,
  messageTimestamp: string | number
): string {
  const timestamp = typeof messageTimestamp === 'string' ? messageTimestamp : String(messageTimestamp);
  return `https://teams.microsoft.com/l/message/${encodeURIComponent(conversationId)}/${timestamp}`;
}

/**
 * Extracts a timestamp-based message ID from various sources.
 * Teams uses epoch milliseconds as message IDs in URLs.
 * 
 * IMPORTANT: For channel threaded replies, the ;messageid= in ClientConversationId
 * is the PARENT thread's ID, not this message's ID. We must prefer the actual
 * message timestamp (DateTimeReceived/DateTimeSent) for accurate deep links.
 */
export function extractMessageTimestamp(
  source: Record<string, unknown> | undefined,
  timestamp?: string
): string | undefined {
  // FIRST: Try to compute from the message's own timestamp
  // This is the most reliable for channel threaded replies
  if (timestamp) {
    try {
      const date = new Date(timestamp);
      if (!isNaN(date.getTime())) {
        return String(date.getTime());
      }
    } catch {
      // Ignore parsing errors
    }
  }
  
  // SECOND: Try explicit MessageId fields
  if (source) {
    // Check for MessageId or Id in various formats
    const messageId = source.MessageId ?? source.OriginalMessageId ?? source.ReferenceObjectId;
    if (typeof messageId === 'string' && /^\d{13}$/.test(messageId)) {
      return messageId;
    }
    
    // LAST RESORT: Check ClientConversationId for ;messageid=xxx suffix
    // NOTE: For threaded replies, this is the PARENT message ID, so only use
    // if we couldn't get the actual timestamp above
    const clientConvId = source.ClientConversationId as string | undefined;
    if (clientConvId && clientConvId.includes(';messageid=')) {
      const match = clientConvId.match(/;messageid=(\d+)/);
      if (match) {
        return match[1];
      }
    }
  }
  
  return undefined;
}

/**
 * Parses a person suggestion from the Substrate API response.
 */
export function parsePersonSuggestion(item: Record<string, unknown>): PersonSearchResult | null {
  const id = item.Id as string;
  if (!id) return null;

  // Extract Azure AD object ID from the Id field (format: "guid@tenantId" or just "guid")
  const objectId = id.split('@')[0];
  
  // Build MRI if not provided
  const mri = (item.MRI as string) || `8:orgid:${objectId}`;
  
  const displayName = item.DisplayName as string || '';
  
  // EmailAddresses can be an array
  const emailAddresses = item.EmailAddresses as string[] | undefined;
  const email = emailAddresses?.[0];

  return {
    id: objectId,
    mri,
    displayName,
    email,
    givenName: item.GivenName as string | undefined,
    surname: item.Surname as string | undefined,
    jobTitle: item.JobTitle as string | undefined,
    department: item.Department as string | undefined,
    companyName: item.CompanyName as string | undefined,
  };
}

/**
 * Parses a v2 query result item into a search result.
 */
export function parseV2Result(item: Record<string, unknown>): TeamsSearchResult | null {
  const content = item.HitHighlightedSummary as string || 
                  item.Summary as string || 
                  '';
  
  if (content.length < 5) return null;

  const id = item.Id as string || 
             item.ReferenceId as string || 
             `v2-${Date.now()}`;

  // Strip HTML from content
  const cleanContent = stripHtml(content);

  const source = item.Source as Record<string, unknown> | undefined;

  // Extract conversationId from extension fields or source properties
  // For channel threaded replies, we want the thread ID (ClientThreadId) not the channel ID
  let conversationId: string | undefined;
  if (source) {
    // Check ClientThreadId first - this is the specific thread for channel replies
    // Using this ensures the deep link goes to the correct thread context
    const clientThreadId = source.ClientThreadId;
    if (typeof clientThreadId === 'string' && clientThreadId.length > 0) {
      conversationId = clientThreadId;
    }
    
    // Fallback to Extensions.SkypeGroupId (the channel ID)
    if (!conversationId) {
      const extensions = source.Extensions as Record<string, unknown> | undefined;
      if (extensions) {
        const extId = extensions.SkypeSpaces_ConversationPost_Extension_SkypeGroupId;
        if (typeof extId === 'string' && extId.length > 0) {
          conversationId = extId;
        }
      }
    }
    
    // Fallback to ClientConversationId (strip ;messageid= suffix if present)
    if (!conversationId) {
      const clientConvId = source.ClientConversationId;
      if (typeof clientConvId === 'string' && clientConvId.length > 0) {
        conversationId = clientConvId.split(';')[0];
      }
    }
  }

  // Note: The API returns DateTimeReceived, DateTimeSent, DateTimeCreated (not ReceivedTime/CreatedDateTime)
  const timestamp = source?.DateTimeReceived as string || 
                    source?.DateTimeSent as string || 
                    source?.DateTimeCreated as string ||
                    source?.ReceivedTime as string ||  // Legacy fallback
                    source?.CreatedDateTime as string; // Legacy fallback
  
  // Build message link if we have the required data
  let messageLink: string | undefined;
  if (conversationId) {
    const messageTimestamp = extractMessageTimestamp(source, timestamp);
    if (messageTimestamp) {
      messageLink = buildMessageLink(conversationId, messageTimestamp);
    }
  }

  return {
    id,
    type: 'message',
    content: cleanContent,
    sender: source?.From as string || source?.Sender as string,
    timestamp,
    channelName: source?.ChannelName as string || source?.Topic as string,
    teamName: source?.TeamName as string || source?.GroupName as string,
    conversationId,
    messageId: item.ReferenceId as string,
    messageLink,
  };
}

/**
 * Parses user profile from a JWT payload.
 * 
 * @param payload - Decoded JWT payload object
 * @returns User profile or null if required fields are missing
 */
export function parseJwtProfile(payload: Record<string, unknown>): UserProfile | null {
  const oid = payload.oid as string | undefined;
  const name = payload.name as string | undefined;
  
  if (!oid || !name) {
    return null;
  }
  
  const profile: UserProfile = {
    id: oid,
    mri: `8:orgid:${oid}`,
    email: (payload.upn || payload.preferred_username || payload.email || '') as string,
    displayName: name,
    tenantId: payload.tid as string | undefined,
  };
  
  // Try to extract given name and surname
  if (payload.given_name) {
    profile.givenName = payload.given_name as string;
  }
  if (payload.family_name) {
    profile.surname = payload.family_name as string;
  }
  
  // If no given/family name, try to parse from displayName
  if (!profile.givenName && profile.displayName.includes(',')) {
    // Format: "Surname, GivenName"
    const parts = profile.displayName.split(',').map(s => s.trim());
    if (parts.length === 2) {
      profile.surname = parts[0];
      profile.givenName = parts[1];
    }
  } else if (!profile.givenName && profile.displayName.includes(' ')) {
    // Format: "GivenName Surname"
    const parts = profile.displayName.split(' ');
    profile.givenName = parts[0];
    profile.surname = parts.slice(1).join(' ');
  }
  
  return profile;
}

/**
 * Calculates token expiry status from an expiry timestamp.
 * 
 * @param expiryMs - Token expiry time in milliseconds since epoch
 * @param nowMs - Current time in milliseconds (for testing)
 * @returns Token status including whether it's valid and time remaining
 */
export function calculateTokenStatus(
  expiryMs: number,
  nowMs: number = Date.now()
): {
  isValid: boolean;
  expiresAt: string;
  minutesRemaining: number;
} {
  const expiryDate = new Date(expiryMs);
  
  return {
    isValid: expiryMs > nowMs,
    expiresAt: expiryDate.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - nowMs) / 1000 / 60)),
  };
}

/**
 * Parses the pagination result from a search API response.
 * 
 * @param entitySets - Raw EntitySets array from API response
 * @param from - Starting offset used in request
 * @param size - Page size used in request
 * @returns Parsed results and pagination metadata
 */
export function parseSearchResults(
  entitySets: unknown[] | undefined,
  from: number,
  size: number
): { results: TeamsSearchResult[]; total?: number } {
  const results: TeamsSearchResult[] = [];
  let total: number | undefined;

  if (!Array.isArray(entitySets)) {
    return { results, total };
  }

  for (const entitySet of entitySets) {
    const es = entitySet as Record<string, unknown>;
    const resultSets = es.ResultSets as unknown[] | undefined;
    
    if (Array.isArray(resultSets)) {
      for (const resultSet of resultSets) {
        const rs = resultSet as Record<string, unknown>;
        
        // Try to get total
        const rsTotal = rs.Total ?? rs.TotalCount ?? rs.TotalEstimate;
        if (typeof rsTotal === 'number') {
          total = rsTotal;
        }
        
        const items = rs.Results as unknown[] | undefined;
        if (Array.isArray(items)) {
          for (const item of items) {
            const parsed = parseV2Result(item as Record<string, unknown>);
            if (parsed) results.push(parsed);
          }
        }
      }
    }
  }

  return { results, total };
}

/**
 * Parses people search results from the Groups/Suggestions structure.
 * 
 * @param groups - Raw Groups array from suggestions API response
 * @returns Array of parsed person results
 */
export function parsePeopleResults(groups: unknown[] | undefined): PersonSearchResult[] {
  const results: PersonSearchResult[] = [];
  
  if (!Array.isArray(groups)) {
    return results;
  }

  for (const group of groups) {
    const g = group as Record<string, unknown>;
    const suggestions = g.Suggestions as unknown[] | undefined;
    
    if (Array.isArray(suggestions)) {
      for (const suggestion of suggestions) {
        const parsed = parsePersonSuggestion(suggestion as Record<string, unknown>);
        if (parsed) results.push(parsed);
      }
    }
  }

  return results;
}
