/**
 * v3 Bootstrap — Singleton Initialization Module
 * 
 * Initializes and exports the v3.0.0 auth components as singletons.
 * Import this module in index.ts to wire up the dual-token system.
 * 
 * Usage in index.ts:
 *   import { v3 } from './v3-bootstrap';
 *   // Then use v3.userAuthManager, v3.tokenRouter, v3.authToolDefinitions
 * 
 * @version 3.0.0
 */

import { UserAuthManager } from './auth/user-auth-manager';
import { TokenRouter } from './auth/token-router';
import { AUTH_TOOL_DEFINITIONS } from './config/auth-tool-definitions';
import { validateDelegatedAuthConfig } from './config/delegated-auth-settings';
import { dispatchAuthTool } from './tools/auth-tool-handlers';

// ——— Singleton Instances ———————————————————————————————————————

let _userAuthManager: UserAuthManager | null = null;
let _tokenRouter: TokenRouter | null = null;

/**
 * Initialize the v3 auth system.
 * Call this once during server startup, AFTER AzureTokenManager is initialized.
 * 
 * @param getServiceToken - Function that returns the service principal token
 *                          (from the existing AzureTokenManager.getToken())
 */
export function initV3Auth(getServiceToken: () => Promise<string>): void {
  console.log('\n[v3] ══════════════════════════════════════════════════');
  console.log('[v3]  Initializing Per-User Delegated Auth (v3.0.0)');
  console.log('[v3] ══════════════════════════════════════════════════\n');

  // Validate config (logs warnings if not configured)
  validateDelegatedAuthConfig();

  // Initialize UserAuthManager
  _userAuthManager = new UserAuthManager();

  // Initialize TokenRouter with both token sources
  _tokenRouter = new TokenRouter({
    getServiceToken,
    userAuthManager: _userAuthManager,
  });

  console.log('\n[v3] ✓ Dual-token system initialized');
  console.log('[v3]   READ  operations → Service Principal');
  console.log('[v3]   WRITE operations → User Delegated Token');
  console.log('[v3] ══════════════════════════════════════════════════\n');
}

/**
 * The v3 module — access all v3 components from a single object.
 */
export const v3 = {
  /** Get the UserAuthManager singleton (throws if not initialized) */
  get userAuthManager(): UserAuthManager {
    if (!_userAuthManager) {
      throw new Error('[v3] Not initialized. Call initV3Auth() during server startup.');
    }
    return _userAuthManager;
  },

  /** Get the TokenRouter singleton (throws if not initialized) */
  get tokenRouter(): TokenRouter {
    if (!_tokenRouter) {
      throw new Error('[v3] Not initialized. Call initV3Auth() during server startup.');
    }
    return _tokenRouter;
  },

  /** Auth tool definitions for MCP tools/list registration */
  authToolDefinitions: AUTH_TOOL_DEFINITIONS,

  /**
   * Dispatch an auth tool call. Returns null if toolName is not an auth tool.
   * Use this in your tools/call handler:
   * 
   *   const authResult = await v3.dispatchAuthTool(toolName, args);
   *   if (authResult) return authResult;
   *   // ... continue with existing tool handling
   */
  async dispatchAuthTool(
    toolName: string,
    args: Record<string, any>
  ): Promise<{ content: Array<{ type: string; text: string }> } | null> {
    return dispatchAuthTool(toolName, args, this.userAuthManager);
  },

  /**
   * Get the appropriate token for a tool call.
   * Use this in PowerPlatformClient or wherever tokens are needed:
   * 
   *   const token = await v3.getToken('pa-create-flow', userId);
   */
  async getToken(toolName: string, userId?: string): Promise<string> {
    return this.tokenRouter.getToken(toolName, userId);
  },

  /** Check if a tool requires user authentication */
  requiresUserAuth(toolName: string): boolean {
    return this.tokenRouter.requiresUserAuth(toolName);
  },
};
