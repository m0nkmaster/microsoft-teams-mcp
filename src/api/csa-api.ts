/**
 * CSA (Chat Service Aggregator) API client for favorites and teams operations.
 * 
 * Handles all calls to teams.microsoft.com/api/csa endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { CSA_API, getCsaHeaders, validateRegion } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractMessageAuth,
  extractCsaToken,
} from '../auth/token-extractor.js';
import {
  getConversationProperties,
  extractParticipantNames,
} from './chatsvc-api.js';
import {
  parseTeamsList,
  type TeamWithChannels,
} from '../utils/parsers.js';

/** A favourite/pinned conversation item. */
export interface FavoriteItem {
  conversationId: string;
  displayName?: string;
  conversationType?: string;
  createdTime?: number;
  lastUpdatedTime?: number;
}

/** Response from getting favorites. */
export interface FavoritesResult {
  favorites: FavoriteItem[];
  folderHierarchyVersion?: number;
  folderId?: string;
}

/**
 * Gets the user's favourite/pinned conversations.
 */
export async function getFavorites(
  region: string = 'amer'
): Promise<Result<FavoritesResult>> {
  const auth = extractMessageAuth();
  const csaToken = extractCsaToken();

  if (!auth?.skypeToken || !csaToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const url = CSA_API.conversationFolders(validRegion);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getCsaHeaders(auth.skypeToken, csaToken),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;

  // Find the Favorites folder
  const folders = data.conversationFolders as unknown[] | undefined;
  const favoritesFolder = folders?.find((f: unknown) => {
    const folder = f as Record<string, unknown>;
    return folder.folderType === 'Favorites';
  }) as Record<string, unknown> | undefined;

  if (!favoritesFolder) {
    return ok({
      favorites: [],
      folderHierarchyVersion: data.folderHierarchyVersion as number,
    });
  }

  const items = favoritesFolder.conversationFolderItems as unknown[] | undefined;
  const favorites: FavoriteItem[] = (items || []).map((item: unknown) => {
    const i = item as Record<string, unknown>;
    return {
      conversationId: i.conversationId as string,
      createdTime: i.createdTime as number | undefined,
      lastUpdatedTime: i.lastUpdatedTime as number | undefined,
    };
  });

  // Enrich favorites with display names in parallel
  const enrichmentPromises = favorites.map(async (fav) => {
    const props = await getConversationProperties(fav.conversationId, validRegion);
    if (props.ok) {
      fav.displayName = props.value.displayName;
      fav.conversationType = props.value.conversationType;
    }

    // Fallback: extract from recent messages if no display name
    if (!fav.displayName) {
      const names = await extractParticipantNames(fav.conversationId, validRegion);
      if (names.ok && names.value) {
        fav.displayName = names.value;
      }
    }
  });

  await Promise.allSettled(enrichmentPromises);

  return ok({
    favorites,
    folderHierarchyVersion: data.folderHierarchyVersion as number,
    folderId: favoritesFolder.id as string,
  });
}

/**
 * Adds a conversation to the user's favourites.
 */
export async function addFavorite(
  conversationId: string,
  region: string = 'amer'
): Promise<Result<void>> {
  return modifyFavorite(conversationId, 'AddItem', region);
}

/**
 * Removes a conversation from the user's favourites.
 */
export async function removeFavorite(
  conversationId: string,
  region: string = 'amer'
): Promise<Result<void>> {
  return modifyFavorite(conversationId, 'RemoveItem', region);
}

/**
 * Internal helper to modify the favourites folder.
 */
async function modifyFavorite(
  conversationId: string,
  action: 'AddItem' | 'RemoveItem',
  region: string
): Promise<Result<void>> {
  const auth = extractMessageAuth();
  const csaToken = extractCsaToken();

  if (!auth?.skypeToken || !csaToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);

  // Get current folder state
  const currentState = await getFavorites(validRegion);
  if (!currentState.ok) {
    return err(currentState.error);
  }

  if (!currentState.value.folderId) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'Could not find Favorites folder'
    ));
  }

  const url = CSA_API.conversationFolders(validRegion);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'POST',
      headers: getCsaHeaders(auth.skypeToken, csaToken),
      body: JSON.stringify({
        folderHierarchyVersion: currentState.value.folderHierarchyVersion,
        actions: [
          {
            action,
            folderId: currentState.value.folderId,
            itemId: conversationId,
          },
        ],
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok(undefined);
}

/** Response from getting the user's teams and channels. */
export interface TeamsListResult {
  teams: TeamWithChannels[];
}

/**
 * Gets all teams and channels the user is a member of.
 * 
 * This returns the complete list of teams with their channels - not a search,
 * but a full enumeration of the user's memberships.
 */
export async function getMyTeamsAndChannels(
  region: string = 'amer'
): Promise<Result<TeamsListResult>> {
  const auth = extractMessageAuth();
  const csaToken = extractCsaToken();

  if (!auth?.skypeToken || !csaToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No valid authentication. Browser login required.'
    ));
  }

  const validRegion = validateRegion(region);
  const url = CSA_API.teamsList(validRegion);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getCsaHeaders(auth.skypeToken, csaToken),
    }
  );

  if (!response.ok) {
    return response;
  }

  const teams = parseTeamsList(response.value.data);

  return ok({ teams });
}
