/**
 * Authentication guard utilities.
 * 
 * Provides reusable auth checks that return Result types for consistent
 * error handling across API modules.
 */

import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { type Result, err, ok } from '../types/result.js';
import {
  getValidSubstrateToken,
  extractMessageAuth,
  extractCsaToken,
  type MessageAuthInfo,
} from '../auth/token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Error Messages
// ─────────────────────────────────────────────────────────────────────────────

const AUTH_ERROR_MESSAGES = {
  substrateToken: 'No valid token available. Browser login required.',
  messageAuth: 'No valid authentication. Browser login required.',
  csaToken: 'No valid authentication for favourites. Browser login required.',
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Guard Types
// ─────────────────────────────────────────────────────────────────────────────

/** Authentication info for messaging and CSA APIs. */
export interface CsaAuthInfo {
  auth: MessageAuthInfo;
  csaToken: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Guard Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Requires a valid Substrate token.
 * Use for search and people APIs.
 */
export function requireSubstrateToken(): Result<string, McpError> {
  const token = getValidSubstrateToken();
  if (!token) {
    return err(createError(ErrorCode.AUTH_REQUIRED, AUTH_ERROR_MESSAGES.substrateToken));
  }
  return ok(token);
}

/**
 * Requires valid message authentication.
 * Use for chatsvc messaging APIs.
 */
export function requireMessageAuth(): Result<MessageAuthInfo, McpError> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(ErrorCode.AUTH_REQUIRED, AUTH_ERROR_MESSAGES.messageAuth));
  }
  return ok(auth);
}

/**
 * Requires valid CSA authentication (message auth + CSA token).
 * Use for favourites and team list APIs.
 */
export function requireCsaAuth(): Result<CsaAuthInfo, McpError> {
  const auth = extractMessageAuth();
  const csaToken = extractCsaToken();

  if (!auth?.skypeToken || !csaToken) {
    return err(createError(ErrorCode.AUTH_REQUIRED, AUTH_ERROR_MESSAGES.csaToken));
  }

  return ok({ auth, csaToken });
}
