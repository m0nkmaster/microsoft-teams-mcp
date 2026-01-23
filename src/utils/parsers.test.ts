/**
 * Unit tests for parsing functions.
 * 
 * Tests outcomes, not implementations - verify that given inputs
 * produce expected outputs regardless of internal logic.
 */

import { describe, it, expect } from 'vitest';
import {
  stripHtml,
  buildMessageLink,
  extractMessageTimestamp,
  parsePersonSuggestion,
  parseV2Result,
  parseJwtProfile,
  calculateTokenStatus,
  parseSearchResults,
  parsePeopleResults,
} from './parsers.js';
import {
  searchResultItem,
  searchResultWithHtml,
  searchResultMinimal,
  searchResultTooShort,
  searchEntitySetsResponse,
  personSuggestion,
  personMinimal,
  peopleGroupsResponse,
  jwtPayloadFull,
  jwtPayloadMinimal,
  jwtPayloadCommaName,
  jwtPayloadSpaceName,
  sourceWithMessageId,
  sourceWithConvIdMessageId,
} from '../__fixtures__/api-responses.js';

describe('stripHtml', () => {
  it('removes HTML tags', () => {
    expect(stripHtml('<p>Hello</p>')).toBe('Hello');
    expect(stripHtml('<div><strong>Bold</strong> text</div>')).toBe('Bold text');
  });

  it('decodes HTML entities', () => {
    expect(stripHtml('Tom &amp; Jerry')).toBe('Tom & Jerry');
    expect(stripHtml('1 &lt; 2 &gt; 0')).toBe('1 < 2 > 0');
    expect(stripHtml('&quot;quoted&quot;')).toBe('"quoted"');
    expect(stripHtml("it&#39;s")).toBe("it's");
    expect(stripHtml('non&nbsp;breaking')).toBe('non breaking');
  });

  it('collapses whitespace', () => {
    expect(stripHtml('hello    world')).toBe('hello world');
    expect(stripHtml('  trimmed  ')).toBe('trimmed');
    expect(stripHtml('line\n\nbreak')).toBe('line break');
  });

  it('handles complex HTML', () => {
    const html = '<p>Meeting <strong>notes</strong> from &amp; yesterday&apos;s call</p><br/><div>Action items:</div>';
    expect(stripHtml(html)).toBe("Meeting notes from & yesterday's call Action items:");
  });

  it('returns empty string for empty input', () => {
    expect(stripHtml('')).toBe('');
  });
});

describe('buildMessageLink', () => {
  it('builds correct Teams deep link', () => {
    const link = buildMessageLink('19:abc@thread.tacv2', '1705760000000');
    expect(link).toBe('https://teams.microsoft.com/l/message/19%3Aabc%40thread.tacv2/1705760000000');
  });

  it('accepts numeric timestamp', () => {
    const link = buildMessageLink('19:abc@thread.v2', 1705760000000);
    expect(link).toBe('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1705760000000');
  });

  it('encodes special characters in conversation ID', () => {
    const link = buildMessageLink('19:special@thread', '123');
    expect(link).toContain('19%3Aspecial%40thread');
  });
});

describe('extractMessageTimestamp', () => {
  it('extracts from MessageId field', () => {
    expect(extractMessageTimestamp(sourceWithMessageId)).toBe('1705760000000');
  });

  it('extracts from ClientConversationId suffix', () => {
    expect(extractMessageTimestamp(sourceWithConvIdMessageId)).toBe('1705770000000');
  });

  it('falls back to parsing ISO timestamp', () => {
    const timestamp = extractMessageTimestamp(undefined, '2026-01-20T12:00:00.000Z');
    expect(timestamp).toBe(String(new Date('2026-01-20T12:00:00.000Z').getTime()));
  });

  it('returns undefined for missing data', () => {
    expect(extractMessageTimestamp(undefined)).toBeUndefined();
    expect(extractMessageTimestamp({})).toBeUndefined();
  });

  it('ignores invalid timestamp formats', () => {
    expect(extractMessageTimestamp(undefined, 'not-a-date')).toBeUndefined();
  });
});

describe('parsePersonSuggestion', () => {
  it('parses complete person data', () => {
    const result = parsePersonSuggestion(personSuggestion);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('a1b2c3d4-e5f6-7890-abcd-ef1234567890');
    expect(result!.mri).toBe('8:orgid:a1b2c3d4-e5f6-7890-abcd-ef1234567890');
    expect(result!.displayName).toBe('Smith, John');
    expect(result!.givenName).toBe('John');
    expect(result!.surname).toBe('Smith');
    expect(result!.email).toBe('john.smith@company.com');
    expect(result!.department).toBe('Engineering');
    expect(result!.jobTitle).toBe('Senior Engineer');
    expect(result!.companyName).toBe('Acme Corp');
  });

  it('handles minimal person data', () => {
    const result = parsePersonSuggestion(personMinimal);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('minimal-user-guid');
    expect(result!.mri).toBe('8:orgid:minimal-user-guid');
    expect(result!.displayName).toBe('Jane Doe');
    expect(result!.email).toBeUndefined();
  });

  it('extracts ID from tenant-qualified format', () => {
    const result = parsePersonSuggestion({
      Id: 'guid123@tenant.onmicrosoft.com',
      DisplayName: 'Test User',
    });
    
    expect(result!.id).toBe('guid123');
  });

  it('returns null for missing ID', () => {
    expect(parsePersonSuggestion({ DisplayName: 'No ID' })).toBeNull();
  });
});

describe('parseV2Result', () => {
  it('parses complete search result', () => {
    const result = parseV2Result(searchResultItem);
    
    expect(result).not.toBeNull();
    expect(result!.type).toBe('message');
    expect(result!.content).toBe('Let me check the budget report for Q3');
    expect(result!.timestamp).toBe('2026-01-20T14:30:00.000Z');
    expect(result!.channelName).toBe('General');
    expect(result!.teamName).toBe('Finance Team');
    expect(result!.conversationId).toBe('19:abcdef123456@thread.tacv2');
    expect(result!.messageLink).toContain('teams.microsoft.com/l/message');
  });

  it('strips HTML from content', () => {
    const result = parseV2Result(searchResultWithHtml);
    
    expect(result).not.toBeNull();
    expect(result!.content).toBe("Meeting notes from & yesterday's call Action items:");
    expect(result!.content).not.toContain('<');
    expect(result!.content).not.toContain('>');
  });

  it('handles minimal result', () => {
    const result = parseV2Result(searchResultMinimal);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('minimal-id');
    expect(result!.content).toBe('A short message here');
    expect(result!.conversationId).toBeUndefined();
    expect(result!.messageLink).toBeUndefined();
  });

  it('returns null for content too short', () => {
    expect(parseV2Result(searchResultTooShort)).toBeNull();
  });

  it('extracts conversationId from Extensions', () => {
    const result = parseV2Result(searchResultItem);
    expect(result!.conversationId).toBe('19:abcdef123456@thread.tacv2');
  });

  it('falls back to ClientThreadId for conversationId', () => {
    const result = parseV2Result(searchResultWithHtml);
    expect(result!.conversationId).toBe('19:meeting123@thread.v2');
  });
});

describe('parseJwtProfile', () => {
  it('parses complete JWT payload', () => {
    const profile = parseJwtProfile(jwtPayloadFull);
    
    expect(profile).not.toBeNull();
    expect(profile!.id).toBe('user-object-id-guid');
    expect(profile!.mri).toBe('8:orgid:user-object-id-guid');
    expect(profile!.email).toBe('rob.macdonald@company.com');
    expect(profile!.displayName).toBe('Macdonald, Rob');
    expect(profile!.givenName).toBe('Rob');
    expect(profile!.surname).toBe('Macdonald');
    expect(profile!.tenantId).toBe('tenant-id-guid');
  });

  it('handles minimal JWT payload', () => {
    const profile = parseJwtProfile(jwtPayloadMinimal);
    
    expect(profile).not.toBeNull();
    expect(profile!.id).toBe('another-user-guid');
    expect(profile!.displayName).toBe('Alice Smith');
    expect(profile!.email).toBe('');
    // Should parse from "Alice Smith" format
    expect(profile!.givenName).toBe('Alice');
    expect(profile!.surname).toBe('Smith');
  });

  it('parses "Surname, GivenName" format', () => {
    const profile = parseJwtProfile(jwtPayloadCommaName);
    
    expect(profile!.surname).toBe('Jones');
    expect(profile!.givenName).toBe('David');
  });

  it('parses "GivenName Surname" format', () => {
    const profile = parseJwtProfile(jwtPayloadSpaceName);
    
    expect(profile!.givenName).toBe('Sarah');
    expect(profile!.surname).toBe('Connor');
  });

  it('returns null for missing required fields', () => {
    expect(parseJwtProfile({})).toBeNull();
    expect(parseJwtProfile({ oid: 'id-only' })).toBeNull();
    expect(parseJwtProfile({ name: 'name-only' })).toBeNull();
  });

  it('prefers upn over other email fields', () => {
    const profile = parseJwtProfile(jwtPayloadFull);
    expect(profile!.email).toBe('rob.macdonald@company.com');
  });
});

describe('calculateTokenStatus', () => {
  const now = 1705846400000; // Fixed "now" for testing

  it('returns valid for unexpired token', () => {
    const expiry = now + 3600000; // 1 hour from now
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.isValid).toBe(true);
    expect(status.minutesRemaining).toBe(60);
  });

  it('returns invalid for expired token', () => {
    const expiry = now - 60000; // 1 minute ago
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.isValid).toBe(false);
    expect(status.minutesRemaining).toBe(0);
  });

  it('returns correct ISO date string', () => {
    const expiry = now + 3600000;
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.expiresAt).toBe(new Date(expiry).toISOString());
  });

  it('rounds minutes correctly', () => {
    const status = calculateTokenStatus(now + 90000, now); // 1.5 minutes
    expect(status.minutesRemaining).toBe(2); // Rounds up
  });
});

describe('parseSearchResults', () => {
  it('parses EntitySets structure', () => {
    const { results, total } = parseSearchResults(
      searchEntitySetsResponse.EntitySets,
      0,
      25
    );
    
    expect(results).toHaveLength(2);
    expect(total).toBe(4307);
  });

  it('returns empty for undefined input', () => {
    const { results, total } = parseSearchResults(undefined, 0, 25);
    
    expect(results).toHaveLength(0);
    expect(total).toBeUndefined();
  });

  it('returns empty for non-array input', () => {
    const { results } = parseSearchResults(
      'not an array' as unknown as unknown[],
      0,
      25
    );
    
    expect(results).toHaveLength(0);
  });

  it('filters out results with short content', () => {
    const entitySets = [{
      ResultSets: [{
        Results: [
          { Id: '1', HitHighlightedSummary: 'Valid content here' },
          { Id: '2', HitHighlightedSummary: 'Hi' }, // Too short
        ],
      }],
    }];
    
    const { results } = parseSearchResults(entitySets, 0, 25);
    expect(results).toHaveLength(1);
  });
});

describe('parsePeopleResults', () => {
  it('parses Groups/Suggestions structure', () => {
    const results = parsePeopleResults(peopleGroupsResponse.Groups);
    
    expect(results).toHaveLength(2);
    expect(results[0].displayName).toBe('Smith, John');
    expect(results[1].displayName).toBe('Jane Doe');
  });

  it('returns empty for undefined input', () => {
    expect(parsePeopleResults(undefined)).toHaveLength(0);
  });

  it('returns empty for non-array input', () => {
    expect(parsePeopleResults('not an array' as unknown as unknown[])).toHaveLength(0);
  });

  it('handles groups with no suggestions', () => {
    const groups = [{ Suggestions: [] }, { OtherField: 'value' }];
    expect(parsePeopleResults(groups)).toHaveLength(0);
  });
});
