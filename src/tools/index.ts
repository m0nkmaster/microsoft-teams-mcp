/**
 * Tool handler registry.
 * 
 * Provides a modular way to define MCP tool handlers without a monolithic
 * switch statement. Each tool is defined with its schema, handler, and metadata.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { z } from 'zod';
import type { McpError } from '../types/errors.js';

import type { BrowserManager } from '../browser/context.js';

// Forward declaration to avoid circular dependency
// TeamsServer is imported dynamically in registry.ts
export interface TeamsServer {
  ensureBrowser(headless?: boolean): Promise<BrowserManager>;
  resetBrowserState(): void;
  getBrowserManager(): BrowserManager | null;
  setBrowserManager(manager: BrowserManager): void;
  markInitialised(): void;
  isInitialisedState(): boolean;
}

/** The context passed to tool handlers. */
export interface ToolContext {
  /** Reference to the server for browser operations. */
  server: TeamsServer;
}

/** Result returned by tool handlers. */
export type ToolResult = 
  | { success: true; data: Record<string, unknown> }
  | { success: false; error: McpError };

/** A registered tool with its handler. */
export interface RegisteredTool<TInput extends z.ZodType = z.ZodType> {
  /** Tool definition for MCP. */
  definition: Tool;
  /** Zod schema for input validation. */
  schema: TInput;
  /** Handler function. */
  handler: (input: z.infer<TInput>, ctx: ToolContext) => Promise<ToolResult>;
}

// Re-export tool registrations
export * from './search-tools.js';
export * from './message-tools.js';
export * from './people-tools.js';
export * from './auth-tools.js';
export * from './registry.js';
