/**
 * TokenRouter — Dual-token dispatcher for Power Automate MCP.
 *
 * Routes API calls to the appropriate token source:
 *   - READ operations → Service Principal (AzureTokenManager)
 *   - WRITE operations → User Delegated Token (UserAuthManager)
 *
 * Phase 2b: Wire this into power-platform-client.ts to enable
 * per-user flow ownership on create/update/delete operations.
 *
 * @version 3.0.1
 */

import { AzureTokenManager } from './azure-token-manager.js';
import { UserAuthManager } from './user-auth-manager.js';

// =============================================================================
// Configuration
// =============================================================================

export interface TokenRouterConfig {
  serviceTokenManager: AzureTokenManager;
  userAuthManager: UserAuthManager;
}

// =============================================================================
// TokenRouter
// =============================================================================

export class TokenRouter {
  private serviceTokenManager: AzureTokenManager;
  private userAuthManager: UserAuthManager;

  constructor(config: TokenRouterConfig) {
    this.serviceTokenManager = config.serviceTokenManager;
    this.userAuthManager = config.userAuthManager;
    console.log('[TokenRouter] Initialized — dual-token routing enabled');
  }

  /**
   * Get the appropriate token for an API operation.
   *
   * @param operation - 'read' uses service principal, 'write' uses user delegated
   * @param userId - Optional user ID for write operations
   * @returns Access token string
   * @throws Error if write operation requested but user not authenticated
   */
  async getToken(operation: 'read' | 'write', userId?: string): Promise<string> {
    if (operation === 'read') {
      // Admin scope — service principal (existing, always available)
      return this.serviceTokenManager.getToken('https://service.flow.microsoft.com/.default');
    }

    // Write operation — needs user's delegated token
    const userToken = await this.userAuthManager.getAccessToken(userId);
    if (!userToken) {
      throw new Error(
        'User authentication required for write operations. ' +
        'Use pa-auth-start to begin device login.'
      );
    }
    return userToken;
  }

  /**
   * Check if a user has a valid token for write operations.
   */
  async canWrite(userId?: string): Promise<boolean> {
    try {
      const token = await this.userAuthManager.getAccessToken(userId);
      return !!token;
    } catch {
      return false;
    }
  }
}
