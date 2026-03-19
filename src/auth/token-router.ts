/**
 * TokenRouter — Dual-Token Dispatcher
 * 
 * Routes authentication based on operation type:
 *   - READ operations  → Service Principal token (AzureTokenManager)
 *   - WRITE operations → User delegated token (UserAuthManager)
 * 
 * This is the key architectural addition for v3.0.0.
 * The Power Interpreter uses pure delegated auth; the Power Automate MCP
 * uses a hybrid model because admin-scoped reads require service principal
 * while flow creation/modification should be user-scoped.
 * 
 * @version 3.0.0
 */

import { UserAuthManager, AuthRequiredError } from './user-auth-manager';

// ─── Types ────────────────────────────────────────────────────────────────────

export type OperationType = 'read' | 'write';

interface TokenRouterConfig {
  /**
   * Get the service principal token (from existing AzureTokenManager).
   * We accept a function rather than the class directly to avoid
   * tight coupling to the existing implementation.
   */
  getServiceToken: () => Promise<string>;
  userAuthManager: UserAuthManager;
}

// ─── Operation Classification ────────────────────────────────────────────────

/**
 * Maps MCP tool names to their operation type.
 * READ tools use service principal; WRITE tools require user auth.
 */
const OPERATION_MAP: Record<string, OperationType> = {
  // ── READ operations (Service Principal) ──
  'pa-list-environments':  'read',
  'pa-list-flows':         'read',
  'pa-get-flow-details':   'read',
  'pa-get-run-history':    'read',
  'pa-get-run-details':    'read',
  'pa-list-connections':   'read',

  // ── WRITE operations (User Delegated) ──
  'pa-create-flow':        'write',
  'pa-update-flow':        'write',
  'pa-delete-flow':        'write',
  'pa-enable-disable-flow':'write',
  'pa-trigger-flow':       'write',
  'pa-cancel-run':         'write',
};

// ─── TokenRouter ─────────────────────────────────────────────────────────────

export class TokenRouter {
  private getServiceToken: () => Promise<string>;
  private userAuthManager: UserAuthManager;

  constructor(config: TokenRouterConfig) {
    this.getServiceToken = config.getServiceToken;
    this.userAuthManager = config.userAuthManager;
    console.log('[TokenRouter] Initialized — dual-token routing active');
  }

  /**
   * Get the appropriate token for the given operation.
   * 
   * @param toolName - The MCP tool being called (e.g., 'pa-list-flows')
   * @param userId - Optional user ID for delegated operations
   * @returns The appropriate access token
   * @throws AuthRequiredError if user auth is needed but not available
   */
  async getToken(toolName: string, userId?: string): Promise<string> {
    const operation = this.classifyOperation(toolName);

    if (operation === 'read') {
      console.log(`[TokenRouter] ${toolName} → service principal (admin read)`);
      return this.getServiceToken();
    }

    // Write operation — needs user token
    console.log(`[TokenRouter] ${toolName} → user delegated token required`);

    if (!this.userAuthManager.isConfigured()) {
      throw new AuthRequiredError(
        'User delegated authentication is not configured. ' +
        'Set PA_USER_CLIENT_ID and PA_USER_TENANT_ID environment variables.'
      );
    }

    const userToken = await this.userAuthManager.getAccessToken(userId);

    if (!userToken) {
      throw new AuthRequiredError(
        'User authentication required for write operations. ' +
        'Use pa-auth-start to begin device login, then pa-auth-poll to complete.'
      );
    }

    return userToken;
  }

  /**
   * Convenience method to get token by operation type directly.
   * Used when the caller already knows the operation type.
   */
  async getTokenByType(operation: OperationType, userId?: string): Promise<string> {
    if (operation === 'read') {
      return this.getServiceToken();
    }

    const userToken = await this.userAuthManager.getAccessToken(userId);
    if (!userToken) {
      throw new AuthRequiredError(
        'User authentication required. Use pa-auth-start to begin device login.'
      );
    }
    return userToken;
  }

  /**
   * Classify a tool name into read or write operation.
   * Unknown tools default to 'read' (safe fallback — service principal).
   */
  private classifyOperation(toolName: string): OperationType {
    const operation = OPERATION_MAP[toolName];
    if (!operation) {
      console.warn(`[TokenRouter] Unknown tool '${toolName}' — defaulting to read (service principal)`);
      return 'read';
    }
    return operation;
  }

  /**
   * Check if a tool requires user authentication.
   * Useful for pre-flight checks before executing a tool.
   */
  requiresUserAuth(toolName: string): boolean {
    return this.classifyOperation(toolName) === 'write';
  }

  /**
   * Get the operation type for a given tool.
   */
  getOperationType(toolName: string): OperationType {
    return this.classifyOperation(toolName);
  }
}
