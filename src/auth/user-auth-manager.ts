/**
 * UserAuthManager — Per-user OAuth 2.0 Device Code Flow
 *
 * Ported from Power Interpreter's MSAuthManager (Python) to TypeScript.
 * Provides delegated authentication for Flow Management API write operations.
 *
 * Architecture:
 *   - Each user authenticates via microsoft.com/devicelogin
 *   - Tokens stored in-memory per userId (email)
 *   - Auto-refresh when < 5 minutes remaining
 *   - Refresh tokens valid for ~90 days
 *   - Multi-scope support: one Device Code auth, refresh token mints tokens
 *     for Flow, PowerApps, Dataverse, or any Microsoft resource scope
 *
 * Integration with PowerPlatformClient:
 *   - getAccessToken(userId?) returns a valid Flow-scoped bearer token
 *   - getAccessTokenForScope(userId?, scope) returns a token for any resource
 *   - Auto-resolves to most recently authenticated user if no userId specified
 *   - Used by createFlow() for delegated write, createConnection() for PowerApps
 *
 * @version 2.1.0
 */

import axios from 'axios';

// =========================================================================
// Exported Types (consumed by auth/index.ts, register-auth-tools.ts, etc.)
// =========================================================================

export interface AuthStartResult {
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  message: string;
}

export interface AuthPollResult {
  status: 'authenticated' | 'pending' | 'expired' | 'error';
  message: string;
  userId?: string;
}

export interface AuthStatusResult {
  authenticated: boolean;
  userId: string | null;
  message: string;
  pending: boolean;
  authenticatedUsers: string[];
  expiresIn?: number;
}

// =========================================================================
// Internal Types
// =========================================================================

interface UserToken {
  accessToken: string;
  refreshToken: string | null;
  expiresAt: number;
  userId: string;
  acquiredAt: number;
}

interface PendingDeviceAuth {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number;
  interval: number;
  userId: string;
}

// =========================================================================
// Constants
// =========================================================================

const REFRESH_BUFFER_MS = 5 * 60 * 1000;  // Refresh when < 5 min remaining
const DEFAULT_SCOPE = 'https://service.flow.microsoft.com/.default offline_access';

// =========================================================================
// UserAuthManager
// =========================================================================

export class UserAuthManager {
  private tenantId: string;
  private clientId: string;
  private scope: string;

  // Primary tokens: userId -> Flow-scoped token (from Device Code auth)
  private tokens: Map<string, UserToken> = new Map();

  // Additional scoped tokens: userId -> (scope -> token)
  // Acquired silently via refresh token exchange
  private scopedTokens: Map<string, Map<string, UserToken>> = new Map();

  // Pending device code auths
  private pendingAuths: Map<string, PendingDeviceAuth> = new Map();

  /**
   * Constructor — accepts optional args; falls back to environment variables.
   * This makes it backward-compatible with existing code that calls new UserAuthManager().
   */
  constructor(tenantId?: string, clientId?: string, scope?: string) {
    this.tenantId = tenantId || process.env.AZURE_TENANT_ID || '';
    this.clientId = clientId || process.env.AZURE_CLIENT_ID || '';
    this.scope = scope || DEFAULT_SCOPE;

    if (this.tenantId && this.clientId) {
      console.log('[UserAuth] Delegated auth configured (Device Code Flow)');
      console.log(`[UserAuth]   Tenant: ${this.tenantId}`);
      console.log(`[UserAuth]   Client: ${this.clientId.substring(0, 8)}...`);
      console.log(`[UserAuth]   Scope: ${this.scope}`);
    } else {
      console.warn('[UserAuth] Missing AZURE_TENANT_ID or AZURE_CLIENT_ID — device code flow disabled');
    }
  }

  // -----------------------------------------------------------------------
  // Azure AD Endpoints
  // -----------------------------------------------------------------------

  private get tokenEndpoint(): string {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
  }

  private get deviceCodeEndpoint(): string {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/devicecode`;
  }

  // -----------------------------------------------------------------------
  // Configuration Check
  // -----------------------------------------------------------------------

  /**
   * Returns true if tenantId and clientId are configured.
   * Used by v3-auth-tools.ts to check if device code flow is available.
   */
  isConfigured(): boolean {
    return !!(this.tenantId && this.clientId);
  }

  // -----------------------------------------------------------------------
  // Device Code Flow: Start
  // -----------------------------------------------------------------------

  /**
   * Initiates the OAuth 2.0 Device Code Flow for a specific user.
   * Returns a user_code and verification_uri for the user to complete sign-in.
   *
   * @param userId - The user's email (e.g., jose@bolthousefresh.com)
   */
  async startAuth(userId: string): Promise<AuthStartResult> {
    if (!this.isConfigured()) {
      throw new Error('Device code flow not configured — AZURE_TENANT_ID and AZURE_CLIENT_ID required.');
    }

    console.log(`[UserAuth] Starting device code flow for: ${userId}`);

    const params = new URLSearchParams({
      client_id: this.clientId,
      scope: this.scope,
    });

    try {
      const response = await axios.post(this.deviceCodeEndpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });

      const data = response.data;

      // Store pending auth state
      this.pendingAuths.set(userId, {
        deviceCode: data.device_code,
        userCode: data.user_code,
        verificationUri: data.verification_uri,
        expiresAt: Date.now() + (data.expires_in * 1000),
        interval: data.interval || 5,
        userId,
      });

      console.log(`[UserAuth] Device code issued for ${userId}: ${data.user_code}`);

      return {
        userCode: data.user_code,
        verificationUri: data.verification_uri,
        expiresIn: data.expires_in,
        message: data.message || `To sign in, visit ${data.verification_uri} and enter code: ${data.user_code}`,
      };
    } catch (error: any) {
      const msg = error.response?.data?.error_description || error.message;
      console.error(`[UserAuth] Device code request failed: ${msg}`);
      throw new Error(`Failed to start authentication: ${msg}`);
    }
  }

  // -----------------------------------------------------------------------
  // Device Code Flow: Poll
  // -----------------------------------------------------------------------

  /**
   * Polls Azure AD to check if the user has completed device code sign-in.
   * Returns 'authenticated' on success, 'pending' while waiting, or 'expired'/'error'.
   *
   * @param userId - The user's email
   */
  async pollAuth(userId: string): Promise<AuthPollResult> {
    const pending = this.pendingAuths.get(userId);
    if (!pending) {
      // Already authenticated?
      if (this.tokens.has(userId)) {
        return {
          status: 'authenticated',
          message: `${userId} is already authenticated.`,
          userId,
        };
      }
      return {
        status: 'error',
        message: `No pending authentication for ${userId}. Use pa-auth-start first.`,
      };
    }

    // Check if device code expired
    if (Date.now() > pending.expiresAt) {
      this.pendingAuths.delete(userId);
      return {
        status: 'expired',
        message: 'Device code expired. Please start a new authentication with pa-auth-start.',
      };
    }

    try {
      const params = new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: this.clientId,
        device_code: pending.deviceCode,
      });

      const response = await axios.post(this.tokenEndpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });

      const data = response.data;

      // Success — store the token
      this.tokens.set(userId, {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || null,
        expiresAt: Date.now() + (data.expires_in * 1000),
        userId,
        acquiredAt: Date.now(),
      });

      this.pendingAuths.delete(userId);

      console.log(`[UserAuth] ✅ ${userId} authenticated successfully`);
      console.log(`[UserAuth]   Token TTL: ${data.expires_in}s, Refresh token: ${data.refresh_token ? 'yes' : 'no'}`);

      return {
        status: 'authenticated',
        message: `Successfully authenticated as ${userId}. You can now create and update flows.`,
        userId,
      };
    } catch (error: any) {
      const errorCode = error.response?.data?.error;
      const errorDesc = error.response?.data?.error_description || '';

      if (errorCode === 'authorization_pending') {
        return {
          status: 'pending',
          message: 'Waiting for user to complete sign-in at microsoft.com/devicelogin...',
        };
      }

      if (errorCode === 'slow_down') {
        return {
          status: 'pending',
          message: 'Polling too fast — please wait a moment and try again.',
        };
      }

      if (errorCode === 'expired_token') {
        this.pendingAuths.delete(userId);
        return {
          status: 'expired',
          message: 'Device code expired. Please start a new authentication with pa-auth-start.',
        };
      }

      console.error(`[UserAuth] Poll error for ${userId}: ${errorCode} — ${errorDesc}`);
      return {
        status: 'error',
        message: `Authentication error: ${errorDesc || errorCode || error.message}`,
      };
    }
  }

  // -----------------------------------------------------------------------
  // Auth Status (dual method names for backward compatibility)
  // -----------------------------------------------------------------------

  /**
   * Returns the current authentication status for a specific user or all users.
   * Aliased as getStatus() for backward compatibility with register-auth-tools.ts
   * and v3-auth-tools.ts.
   */
  getAuthStatus(userId?: string): AuthStatusResult {
    const authenticatedUsers = this.listAuthenticatedUsers();

    if (userId) {
      const token = this.tokens.get(userId);
      const hasPending = this.pendingAuths.has(userId);

      if (token && (token.expiresAt > Date.now() || token.refreshToken)) {
        const secondsLeft = Math.max(0, Math.round((token.expiresAt - Date.now()) / 1000));
        const minutesLeft = Math.max(0, Math.round(secondsLeft / 60));
        return {
          authenticated: true,
          userId,
          message: `${userId} is authenticated (expires in ${minutesLeft} min, refresh: ${token.refreshToken ? 'available' : 'none'})`,
          pending: false,
          authenticatedUsers,
          expiresIn: secondsLeft,
        };
      }

      return {
        authenticated: false,
        userId,
        message: hasPending
          ? `${userId} has a pending device code login — use pa-auth-poll to complete.`
          : `${userId} is not authenticated. Use pa-auth-start to begin.`,
        pending: hasPending,
        authenticatedUsers,
        expiresIn: 0,
      };
    }

    // No specific user — return general status
    // If there's a default user, include their expiresIn
    let expiresIn = 0;
    const defaultUser = this.getDefaultUserId();
    if (defaultUser) {
      const token = this.tokens.get(defaultUser);
      if (token) {
        expiresIn = Math.max(0, Math.round((token.expiresAt - Date.now()) / 1000));
      }
    }

    return {
      authenticated: authenticatedUsers.length > 0,
      userId: authenticatedUsers[0] || null,
      message: authenticatedUsers.length > 0
        ? `${authenticatedUsers.length} user(s) authenticated: ${authenticatedUsers.join(', ')}`
        : 'No users authenticated. Use pa-auth-start to begin device code login.',
      pending: this.pendingAuths.size > 0,
      authenticatedUsers,
      expiresIn,
    };
  }

  /**
   * Alias for getAuthStatus() — called by register-auth-tools.ts and v3-auth-tools.ts.
   */
  getStatus(userId?: string): AuthStatusResult {
    return this.getAuthStatus(userId);
  }

  // -----------------------------------------------------------------------
  // Access Token — Primary (Flow scope)
  // -----------------------------------------------------------------------

  /**
   * Gets a valid access token for the specified user (Flow scope).
   * Auto-refreshes if the token is within 5 minutes of expiry.
   * Returns null if no token is available (user must authenticate).
   *
   * This is the primary integration point with PowerPlatformClient.
   * Called by userFlowRequest() for delegated write operations.
   */
  async getAccessToken(userId?: string): Promise<string | null> {
    const targetUser = userId || this.getDefaultUserId();
    if (!targetUser) {
      console.log('[UserAuth] No authenticated user available for token request');
      return null;
    }

    const token = this.tokens.get(targetUser);
    if (!token) {
      console.log(`[UserAuth] No token found for ${targetUser}`);
      return null;
    }

    // Auto-refresh if < 5 min remaining
    if (token.expiresAt - Date.now() < REFRESH_BUFFER_MS) {
      if (token.refreshToken) {
        console.log(`[UserAuth] Token for ${targetUser} expiring soon, refreshing...`);
        const refreshed = await this.refreshTokenForUser(targetUser, token.refreshToken);
        if (refreshed) {
          return this.tokens.get(targetUser)!.accessToken;
        }
      }

      // Token expired and refresh failed or unavailable
      if (token.expiresAt < Date.now()) {
        console.warn(`[UserAuth] Token expired for ${targetUser} — re-authentication required`);
        this.tokens.delete(targetUser);
        return null;
      }
    }

    return token.accessToken;
  }

  // -----------------------------------------------------------------------
  // Access Token — Multi-Scope (v2.1.0)
  //
  // Uses the stored refresh token to silently acquire access tokens for
  // ANY Microsoft resource scope. The user authenticates ONCE via Device
  // Code Flow; the refresh token mints tokens for additional resources.
  //
  // Example:
  //   getAccessTokenForScope(userId, 'https://service.powerapps.com/.default')
  //   getAccessTokenForScope(userId, 'https://graph.microsoft.com/.default')
  // -----------------------------------------------------------------------

  /**
   * Gets an access token for a specific resource scope.
   * If scope is the default Flow scope (or omitted), delegates to getAccessToken().
   * Otherwise, checks the scoped token cache and acquires via refresh token exchange.
   *
   * @param userId - User email (optional, defaults to most recent authenticated user)
   * @param scope  - Target resource scope (e.g., 'https://service.powerapps.com/.default')
   * @returns Access token string or null if unavailable
   */
  async getAccessTokenForScope(userId?: string, scope?: string): Promise<string | null> {
    // No scope or default Flow scope → use primary token path
    if (!scope || scope === DEFAULT_SCOPE || scope.startsWith('https://service.flow.microsoft.com/')) {
      return this.getAccessToken(userId);
    }

    const targetUser = userId || this.getDefaultUserId();
    if (!targetUser) {
      console.log('[UserAuth] No authenticated user available for scoped token request');
      return null;
    }

    // Check scoped token cache
    const userScoped = this.scopedTokens.get(targetUser);
    if (userScoped) {
      const cached = userScoped.get(scope);
      if (cached) {
        // Auto-refresh scoped token if < 5 min remaining
        if (cached.expiresAt - Date.now() >= REFRESH_BUFFER_MS) {
          return cached.accessToken;
        }
        console.log(`[UserAuth] Scoped token (${this.shortScope(scope)}) for ${targetUser} expiring soon, re-acquiring...`);
      }
    }

    // Acquire new scoped token via refresh token exchange
    const primaryToken = this.tokens.get(targetUser);
    if (!primaryToken?.refreshToken) {
      console.log(`[UserAuth] No refresh token for ${targetUser} — cannot acquire ${this.shortScope(scope)} token`);
      return null;
    }

    return this.acquireScopedToken(targetUser, scope, primaryToken.refreshToken);
  }

  // -----------------------------------------------------------------------
  // Token Refresh — Primary (internal)
  // -----------------------------------------------------------------------

  private async refreshTokenForUser(userId: string, refreshToken: string): Promise<boolean> {
    try {
      console.log(`[UserAuth] Refreshing token for ${userId}...`);

      const params = new URLSearchParams({
        grant_type: 'refresh_token',
        client_id: this.clientId,
        refresh_token: refreshToken,
        scope: this.scope,
      });

      const response = await axios.post(this.tokenEndpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });

      const data = response.data;

      this.tokens.set(userId, {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || refreshToken,
        expiresAt: Date.now() + (data.expires_in * 1000),
        userId,
        acquiredAt: Date.now(),
      });

      console.log(`[UserAuth] ✅ Token refreshed for ${userId} (TTL: ${data.expires_in}s)`);
      return true;
    } catch (error: any) {
      const errMsg = error.response?.data?.error_description || error.response?.data?.error || error.message;
      console.error(`[UserAuth] Refresh failed for ${userId}: ${errMsg}`);
      return false;
    }
  }

  // -----------------------------------------------------------------------
  // Token Acquisition — Scoped (v2.1.0, internal)
  //
  // Exchanges the stored refresh token for an access token targeting a
  // different Microsoft resource. Azure AD v2.0 allows a single refresh
  // token to be used across any resource the app has permissions for.
  // -----------------------------------------------------------------------

  private async acquireScopedToken(
    userId: string,
    scope: string,
    refreshToken: string
  ): Promise<string | null> {
    try {
      // Ensure offline_access is included so we get a fresh refresh token back
      const requestScope = scope.includes('offline_access')
        ? scope
        : `${scope} offline_access`;

      console.log(`[UserAuth] Acquiring ${this.shortScope(scope)} token for ${userId} via refresh token exchange...`);

      const params = new URLSearchParams({
        grant_type: 'refresh_token',
        client_id: this.clientId,
        refresh_token: refreshToken,
        scope: requestScope,
      });

      const response = await axios.post(this.tokenEndpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });

      const data = response.data;

      // Store scoped token
      if (!this.scopedTokens.has(userId)) {
        this.scopedTokens.set(userId, new Map());
      }
      this.scopedTokens.get(userId)!.set(scope, {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || refreshToken,
        expiresAt: Date.now() + (data.expires_in * 1000),
        userId,
        acquiredAt: Date.now(),
      });

      // Sync refresh token back to primary if Azure AD rotated it
      if (data.refresh_token) {
        const primary = this.tokens.get(userId);
        if (primary && data.refresh_token !== primary.refreshToken) {
          primary.refreshToken = data.refresh_token;
          console.log(`[UserAuth] Refresh token rotated for ${userId} (synced to primary)`);
        }
      }

      console.log(`[UserAuth] ✅ ${this.shortScope(scope)} token acquired for ${userId} (TTL: ${data.expires_in}s)`);
      return data.access_token;

    } catch (error: any) {
      const errCode = error.response?.data?.error || '';
      const errMsg = error.response?.data?.error_description || error.message;

      // If the error is interaction_required, the user may need to re-consent
      if (errCode === 'interaction_required' || errCode === 'invalid_grant') {
        console.error(
          `[UserAuth] Cannot silently acquire ${this.shortScope(scope)} token for ${userId}: ${errCode}. ` +
          'The user may need to re-authenticate or admin consent may be required for this scope.'
        );
      } else {
        console.error(`[UserAuth] Scoped token acquisition failed for ${userId} (${this.shortScope(scope)}): ${errMsg}`);
      }

      return null;
    }
  }

  /**
   * Returns a short human-readable label for a scope URL.
   * e.g., 'https://service.powerapps.com/.default' → 'powerapps'
   */
  private shortScope(scope: string): string {
    if (scope.includes('powerapps')) return 'powerapps';
    if (scope.includes('flow')) return 'flow';
    if (scope.includes('graph')) return 'graph';
    if (scope.includes('dynamics') || scope.includes('crm')) return 'dataverse';
    return scope.split('/')[2] || scope;
  }

  // -----------------------------------------------------------------------
  // Utility Methods
  // -----------------------------------------------------------------------

  /**
   * Returns true if any user has a valid (or refreshable) token.
   */
  hasAuthenticatedUser(): boolean {
    for (const [, token] of this.tokens) {
      if (token.expiresAt > Date.now() || token.refreshToken) {
        return true;
      }
    }
    return false;
  }

  /**
   * Returns a list of all currently authenticated user IDs (emails).
   * Called by register-auth-tools.ts and v3-auth-tools.ts.
   */
  listAuthenticatedUsers(): string[] {
    const users: string[] = [];
    for (const [userId, token] of this.tokens) {
      if (token.expiresAt > Date.now() || token.refreshToken) {
        users.push(userId);
      }
    }
    return users;
  }

  /**
   * Returns the userId of the most recently authenticated user.
   */
  getDefaultUserId(): string | null {
    let latest: UserToken | null = null;
    for (const [, token] of this.tokens) {
      if (!latest || token.acquiredAt > latest.acquiredAt) {
        latest = token;
      }
    }
    return latest?.userId || null;
  }

  /**
   * Invalidates a specific user's token (e.g., on 401 retry failure).
   * Also clears all scoped tokens for the user.
   */
  invalidateUser(userId: string): void {
    this.tokens.delete(userId);
    this.scopedTokens.delete(userId);
    console.log(`[UserAuth] Token invalidated for ${userId} (primary + all scoped tokens)`);
  }
}
