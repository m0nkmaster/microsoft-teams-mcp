/**
 * Shared constants used across the codebase.
 * 
 * Centralising these values makes the code more maintainable and
 * allows for easier configuration changes.
 */

// ─────────────────────────────────────────────────────────────────────────────
// Content Thresholds
// ─────────────────────────────────────────────────────────────────────────────

/** Minimum content length to be considered valid (characters). */
export const MIN_CONTENT_LENGTH = 5;

// ─────────────────────────────────────────────────────────────────────────────
// Pagination Defaults
// ─────────────────────────────────────────────────────────────────────────────

/** Default page size for search results. */
export const DEFAULT_PAGE_SIZE = 25;

/** Maximum page size for search results. */
export const MAX_PAGE_SIZE = 100;

/** Default limit for thread messages. */
export const DEFAULT_THREAD_LIMIT = 50;

/** Maximum limit for thread messages. */
export const MAX_THREAD_LIMIT = 200;

/** Default limit for people search. */
export const DEFAULT_PEOPLE_LIMIT = 10;

/** Maximum limit for people search. */
export const MAX_PEOPLE_LIMIT = 50;

/** Default limit for frequent contacts. */
export const DEFAULT_CONTACTS_LIMIT = 50;

/** Maximum limit for frequent contacts. */
export const MAX_CONTACTS_LIMIT = 500;

/** Default limit for channel search. */
export const DEFAULT_CHANNEL_LIMIT = 10;

/** Maximum limit for channel search. */
export const MAX_CHANNEL_LIMIT = 50;

// ─────────────────────────────────────────────────────────────────────────────
// Timeouts (milliseconds)
// ─────────────────────────────────────────────────────────────────────────────

/** Default timeout for waiting for search results. */
export const SEARCH_RESULT_TIMEOUT_MS = 10000;

/** Default HTTP request timeout. */
export const HTTP_REQUEST_TIMEOUT_MS = 30000;

/** Short delay for UI interactions. */
export const UI_SHORT_DELAY_MS = 300;

/** Medium delay for UI state changes. */
export const UI_MEDIUM_DELAY_MS = 1000;

/** Long delay for API responses to settle. */
export const UI_LONG_DELAY_MS = 2000;

/** Authentication check interval. */
export const AUTH_CHECK_INTERVAL_MS = 2000;

/** Default login timeout (5 minutes). */
export const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

/** Pause after showing progress overlay step (ms). */
export const OVERLAY_STEP_PAUSE_MS = 1500;

/** Pause after showing final "All done" overlay (ms). */
export const OVERLAY_COMPLETE_PAUSE_MS = 2000;

// ─────────────────────────────────────────────────────────────────────────────
// Session Management
// ─────────────────────────────────────────────────────────────────────────────

/** Session expiry threshold in hours. */
export const SESSION_EXPIRY_HOURS = 12;

// ─────────────────────────────────────────────────────────────────────────────
// Retry Configuration
// ─────────────────────────────────────────────────────────────────────────────

/** Default maximum retry attempts for HTTP requests. */
export const DEFAULT_MAX_RETRIES = 3;

/** Base delay for exponential backoff (milliseconds). */
export const RETRY_BASE_DELAY_MS = 1000;

/** Maximum delay between retries (milliseconds). */
export const RETRY_MAX_DELAY_MS = 10000;

// ─────────────────────────────────────────────────────────────────────────────
// Conversation IDs
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Virtual Conversation IDs.
 * 
 * These special IDs are used with the standard chatsvc messages endpoint
 * (/users/ME/conversations/{id}/messages) to retrieve aggregated views
 * across all conversations. See docs/API-REFERENCE.md for details.
 */

/** Prefix for virtual conversation IDs (48:saved, 48:notifications, etc). */
export const VIRTUAL_CONVERSATION_PREFIX = '48:';

/** Self-chat (notes) conversation ID. */
export const SELF_CHAT_ID = '48:notes';

/** Activity feed (notifications) conversation ID. */
export const NOTIFICATIONS_ID = '48:notifications';

/** Saved messages virtual conversation ID. */
export const SAVED_MESSAGES_ID = '48:saved';

/** Followed threads virtual conversation ID. */
export const FOLLOWED_THREADS_ID = '48:threads';

// ─────────────────────────────────────────────────────────────────────────────
// Activity Feed
// ─────────────────────────────────────────────────────────────────────────────

/** Default limit for activity feed items. */
export const DEFAULT_ACTIVITY_LIMIT = 50;

/** Maximum limit for activity feed items. */
export const MAX_ACTIVITY_LIMIT = 200;

// ─────────────────────────────────────────────────────────────────────────────
// Unread Status
// ─────────────────────────────────────────────────────────────────────────────

/** Maximum conversations to check when aggregating unread status. */
export const MAX_UNREAD_AGGREGATE_CHECK = 20;

// ─────────────────────────────────────────────────────────────────────────────
// Token Refresh
// ─────────────────────────────────────────────────────────────────────────────

/** Threshold for proactive token refresh (10 minutes before expiry). */
export const TOKEN_REFRESH_THRESHOLD_MS = 10 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// User Identity
// ─────────────────────────────────────────────────────────────────────────────

/** MRI prefix for organisation users (orgid). */
export const MRI_ORGID_PREFIX = '8:orgid:';
