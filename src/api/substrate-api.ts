/**
 * Substrate API client for search and people operations.
 * 
 * Handles all calls to substrate.office.com endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { SUBSTRATE_API, getBearerHeaders } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  getValidSubstrateToken,
  clearTokenCache,
} from '../auth/token-extractor.js';
import {
  parseSearchResults,
  parsePeopleResults,
  type PersonSearchResult,
} from '../utils/parsers.js';
import type { TeamsSearchResult, SearchPaginationResult } from '../types/teams.js';

/** Search result with pagination. */
export interface SearchResult {
  results: TeamsSearchResult[];
  pagination: SearchPaginationResult;
}

/** People search result. */
export interface PeopleSearchResult {
  results: PersonSearchResult[];
  returned: number;
}

/**
 * Searches Teams messages using the Substrate v2 query API.
 */
export async function searchMessages(
  query: string,
  options: { from?: number; size?: number; maxResults?: number } = {}
): Promise<Result<SearchResult>> {
  const token = getValidSubstrateToken();
  if (!token) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid token available. Browser login required.'
    ));
  }

  const from = options.from ?? 0;
  const size = options.size ?? 25;

  // Generate unique IDs for this request
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    entityRequests: [{
      entityType: 'Message',
      contentSources: ['Teams'],
      propertySet: 'Optimized',
      fields: [
        'Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String',
      ],
      query: {
        queryString: `${query} AND NOT (isClientSoftDeleted:TRUE)`,
        displayQueryString: query,
      },
      from,
      size,
      topResultsCount: 5,
    }],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ['Suggestion'],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Messages' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
      Name: 'powerbar',
    },
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.search,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    // Clear cache on auth errors
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const data = response.value.data;
  const { results, total } = parseSearchResults(
    data.EntitySets as unknown[] | undefined,
    from,
    size
  );

  const maxResults = options.maxResults ?? size;
  const limitedResults = results.slice(0, maxResults);

  return ok({
    results: limitedResults,
    pagination: {
      from,
      size,
      returned: limitedResults.length,
      total,
      hasMore: total !== undefined
        ? from + results.length < total
        : results.length >= size,
    },
  });
}

/**
 * Searches for people by name or email.
 */
export async function searchPeople(
  query: string,
  limit: number = 10
): Promise<Result<PeopleSearchResult>> {
  const token = getValidSubstrateToken();
  if (!token) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid token available. Browser login required.'
    ));
  }

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.peopleSearch,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const results = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results,
    returned: results.length,
  });
}

/**
 * Gets the user's frequently contacted people.
 */
export async function getFrequentContacts(
  limit: number = 50
): Promise<Result<PeopleSearchResult>> {
  const token = getValidSubstrateToken();
  if (!token) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid token available. Browser login required.'
    ));
  }

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: '',
        DisplayQueryString: '',
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.frequentContacts,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const contacts = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results: contacts,
    returned: contacts.length,
  });
}
