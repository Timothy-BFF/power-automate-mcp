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
 *
 * Integration with PowerPlatformClient:
 *   - getAccessToken(userId?) returns a valid bearer token
 *   - Auto-resolves to most recently authenticated user if no userId specified
 *   - Used by createFlow() and updateFlow() for delegated write operations
 *
 * @version 2.0.0
 */

import axios from 'axios';

// =========================================================================
// Types
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
  private tokens: Map<string, UserToken> = new Map();
  private pendingAuths: Map<string, PendingDeviceAuth> = new Map();

  constructor(tenantId: string, clientId: string, scope?: string) {
    this.tenantId = tenantId;
    this.clientId = clientId;
    this.scope = scope || DEFAULT_SCOPE;

    console.log('[UserAuth] Delegated auth configured (Device Code Flow)');
    console.log(`[UserAuth]   Tenant: ${tenantId}`);
    console.log(`[UserAuth]   Client: ${clientId.substring(0, 8)}...`);
    console.log(`[UserAuth]   Scope: ${this.scope}`);
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
  // Device Code Flow: Start
  // -----------------------------------------------------------------------

  /**
   * Initiates the OAuth 2.0 Device Code Flow for a specific user.
   * Returns a user_code and verification_uri for the user to complete sign-in.
   *
   * @param userId - The user's email (e.g., jose@bolthousefresh.com)
   */
  async startAuth(userId: string): Promise<{
    userCode: string;
    verificationUri: string;
    expiresIn: number;
    message: string;
  }> {
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
  async pollAuth(userId: string): Promise<{
    status: 'authenticated' | 'pending' | 'expired' | 'error';
    message: string;
    userId?: string;
  }> {
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

      console.log(`[UserAuth] \u2705 ${userId} authenticated successfully`);
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
          message: 'Polling too fast \u2014 please wait a moment and try again.',
        };
      }

      if (errorCode === 'expired_token') {
        this.pendingAuths.delete(userId);
        return {
          status: 'expired',
          message: 'Device code expired. Please start a new authentication with pa-auth-start.',
        };
      }

      console.error(`[UserAuth] Poll error for ${userId}: ${errorCode} \u2014 ${errorDesc}`);
      return {
        status: 'error',
        message: `Authentication error: ${errorDesc || errorCode || error.message}`,
      };
    }
  }

  // -----------------------------------------------------------------------
  // Auth Status
  // -----------------------------------------------------------------------

  /**
   * Returns the current authentication status for a specific user or all users.
   */
  getAuthStatus(userId?: string): {
    authenticated: boolean;
    userId: string | null;
    message: string;
    pending: boolean;
    authenticatedUsers: string[];
  } {
    const authenticatedUsers = Array.from(this.tokens.keys()).filter(uid => {
      const t = this.tokens.get(uid)!;
      return t.expiresAt > Date.now() || t.refreshToken;
    });

    if (userId) {
      const token = this.tokens.get(userId);
      const hasPending = this.pendingAuths.has(userId);

      if (token && (token.expiresAt > Date.now() || token.refreshToken)) {
        const minutesLeft = Math.max(0, Math.round((token.expiresAt - Date.now()) / 60000));
        return {
          authenticated: true,
          userId,
          message: `${userId} is authenticated (expires in ${minutesLeft} min, refresh: ${token.refreshToken ? 'available' : 'none'})`,
          pending: false,
          authenticatedUsers,
        };
      }

      return {
        authenticated: false,
        userId,
        message: hasPending
          ? `${userId} has a pending device code login \u2014 use pa-auth-poll to complete.`
          : `${userId} is not authenticated. Use pa-auth-start to begin.`,
        pending: hasPending,
        authenticatedUsers,
      };
    }

    // No specific user \u2014 return general status
    return {
      authenticated: authenticatedUsers.length > 0,
      userId: authenticatedUsers[0] || null,
      message: authenticatedUsers.length > 0
        ? `${authenticatedUsers.length} user(s) authenticated: ${authenticatedUsers.join(', ')}`
        : 'No users authenticated. Use pa-auth-start to begin device code login.',
      pending: this.pendingAuths.size > 0,
      authenticatedUsers,
    };
  }

  // -----------------------------------------------------------------------
  // Access Token (for PowerPlatformClient integration)
  // -----------------------------------------------------------------------

  /**
   * Gets a valid access token for the specified user.
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
        console.warn(`[UserAuth] Token expired for ${targetUser} \u2014 re-authentication required`);
        this.tokens.delete(targetUser);
        return null;
      }
    }

    return token.accessToken;
  }

  // -----------------------------------------------------------------------
  // Token Refresh (internal)
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

      console.log(`[UserAuth] \u2705 Token refreshed for ${userId} (TTL: ${data.expires_in}s)`);
      return true;
    } catch (error: any) {
      const errMsg = error.response?.data?.error_description || error.response?.data?.error || error.message;
      console.error(`[UserAuth] Refresh failed for ${userId}: ${errMsg}`);
      return false;
    }
  }

  // -----------------------------------------------------------------------
  // Utility
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
   */
  invalidateUser(userId: string): void {
    this.tokens.delete(userId);
    console.log(`[UserAuth] Token invalidated for ${userId}`);
  }
}
