/**
 * UserAuthManager — Per-user delegated authentication via OAuth 2.0 Device Code Flow.
 *
 * Ported from Power Interpreter MCP's MSAuthManager (Python) to TypeScript.
 * Handles the full device code lifecycle: initiate → poll → token store → auto-refresh.
 *
 * Token storage: In-memory (survives for the lifetime of the process).
 * Future: PostgreSQL persistence (same pattern as Power Interpreter).
 *
 * @version 3.0.1
 */

import axios from 'axios';

// =============================================================================
// Interfaces
// =============================================================================

/** Mapped from Azure AD device code endpoint response (snake_case → camelCase) */
interface DeviceCodeResponse {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  interval: number;
  message: string;
}

/** Mapped from Azure AD token endpoint response (snake_case → camelCase) */
interface TokenResponse {
  accessToken: string;
  refreshToken?: string;
  expiresIn: number;
  tokenType: string;
  scope?: string;
}

/** Per-user token data stored in memory */
interface UserTokenData {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number; // Unix timestamp (ms)
  userId: string;
  acquiredAt: number; // Unix timestamp (ms)
}

/** Tracks a pending device code auth flow */
interface PendingAuth {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number; // Unix timestamp (ms)
  interval: number;
  userId: string;
}

/** Result returned by startAuth */
export interface AuthStartResult {
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  message: string;
}

/** Result returned by pollAuth */
export interface AuthPollResult {
  status: 'authenticated' | 'pending' | 'expired' | 'error';
  message: string;
  userId?: string;
}

/** Result returned by getStatus */
export interface AuthStatusResult {
  authenticated: boolean;
  message: string;
  userId?: string;
  expiresIn?: number;
}

// =============================================================================
// Constants
// =============================================================================

const FLOW_SCOPE = 'https://service.flow.microsoft.com/.default offline_access';
const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000; // Refresh if < 5 min remaining

// =============================================================================
// UserAuthManager
// =============================================================================

export class UserAuthManager {
  private tokens: Map<string, UserTokenData> = new Map();
  private pendingAuths: Map<string, PendingAuth> = new Map();
  private clientId: string;
  private tenantId: string;

  constructor() {
    this.clientId = process.env.PA_USER_CLIENT_ID || '';
    this.tenantId = process.env.PA_USER_TENANT_ID || process.env.AZURE_TENANT_ID || '';

    if (this.clientId && this.tenantId) {
      console.log('[UserAuth] Delegated auth configured (Device Code Flow)');
      console.log(`[UserAuth]   Tenant: ${this.tenantId}`);
      console.log(`[UserAuth]   Client: ${this.clientId.substring(0, 8)}...`);
      console.log(`[UserAuth]   Scope:  ${FLOW_SCOPE}`);
    } else {
      console.warn('[UserAuth] Delegated auth NOT configured — missing PA_USER_CLIENT_ID or PA_USER_TENANT_ID');
      console.warn('[UserAuth] Auth tools will be registered but will return configuration errors');
    }
  }

  // ===========================================================================
  // Public: Configuration Check
  // ===========================================================================

  /** Returns true if the required env vars are set */
  isConfigured(): boolean {
    return !!(this.clientId && this.tenantId);
  }

  // ===========================================================================
  // Public: Start Device Code Flow
  // ===========================================================================

  async startAuth(userId: string): Promise<AuthStartResult> {
    if (!this.isConfigured()) {
      throw new Error(
        'Delegated auth not configured. Set PA_USER_CLIENT_ID and PA_USER_TENANT_ID environment variables.'
      );
    }

    const normalizedId = userId.toLowerCase().trim();
    console.log(`[UserAuth] Starting device code flow for: ${normalizedId}`);

    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/devicecode`;

    const response = await axios.post(
      url,
      new URLSearchParams({
        client_id: this.clientId,
        scope: FLOW_SCOPE,
      }).toString(),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );

    // Azure AD returns snake_case — map to our camelCase interface
    const raw = response.data;
    const deviceCode: DeviceCodeResponse = {
      deviceCode: raw.device_code,
      userCode: raw.user_code,
      verificationUri: raw.verification_uri,
      expiresIn: raw.expires_in,
      interval: raw.interval || 5,
      message: raw.message,
    };

    // Store pending auth
    this.pendingAuths.set(normalizedId, {
      deviceCode: deviceCode.deviceCode,
      userCode: deviceCode.userCode,
      verificationUri: deviceCode.verificationUri,
      expiresAt: Date.now() + deviceCode.expiresIn * 1000,
      interval: deviceCode.interval,
      userId: normalizedId,
    });

    console.log(`[UserAuth] Device code issued for ${normalizedId}: ${deviceCode.userCode}`);
    console.log(`[UserAuth] Verification URL: ${deviceCode.verificationUri}`);
    console.log(`[UserAuth] Expires in: ${deviceCode.expiresIn}s`);

    return {
      userCode: deviceCode.userCode,
      verificationUri: deviceCode.verificationUri,
      expiresIn: deviceCode.expiresIn,
      message: deviceCode.message,
    };
  }

  // ===========================================================================
  // Public: Poll for Auth Completion
  // ===========================================================================

  async pollAuth(userId?: string): Promise<AuthPollResult> {
    if (!this.isConfigured()) {
      throw new Error('Delegated auth not configured.');
    }

    const resolvedId = this.resolvePendingUserId(userId);
    if (!resolvedId) {
      return {
        status: 'error',
        message: userId
          ? `No pending authentication for user: ${userId}`
          : 'No pending authentication found. Call pa-auth-start first.',
      };
    }

    const pending = this.pendingAuths.get(resolvedId)!;

    // Check expiry
    if (Date.now() > pending.expiresAt) {
      this.pendingAuths.delete(resolvedId);
      return {
        status: 'expired',
        message: 'Device code expired. Call pa-auth-start to begin a new authentication.',
        userId: resolvedId,
      };
    }

    // Poll Azure AD token endpoint
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

    try {
      const response = await axios.post(
        url,
        new URLSearchParams({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: this.clientId,
          device_code: pending.deviceCode,
        }).toString(),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          validateStatus: () => true, // Don't throw on 400 (pending)
        }
      );

      if (response.status === 200) {
        // Azure AD returns snake_case — map to camelCase
        const raw = response.data;
        const tokenData: TokenResponse = {
          accessToken: raw.access_token,
          refreshToken: raw.refresh_token,
          expiresIn: raw.expires_in,
          tokenType: raw.token_type,
          scope: raw.scope,
        };

        // Store token
        this.tokens.set(resolvedId, {
          accessToken: tokenData.accessToken,
          refreshToken: tokenData.refreshToken,
          expiresAt: Date.now() + tokenData.expiresIn * 1000,
          userId: resolvedId,
          acquiredAt: Date.now(),
        });

        // Clean up pending
        this.pendingAuths.delete(resolvedId);

        console.log(`[UserAuth] ✅ Authentication complete for: ${resolvedId}`);
        console.log(`[UserAuth]   Token expires in: ${tokenData.expiresIn}s`);
        console.log(`[UserAuth]   Refresh token: ${tokenData.refreshToken ? 'present' : 'none'}`);

        return {
          status: 'authenticated',
          message: `Successfully authenticated as ${resolvedId}. Write operations are now available.`,
          userId: resolvedId,
        };
      }

      // 400 with authorization_pending = user hasn't completed login yet
      if (response.data?.error === 'authorization_pending') {
        return {
          status: 'pending',
          message: `Waiting for user to complete login at ${pending.verificationUri} with code ${pending.userCode}`,
          userId: resolvedId,
        };
      }

      // 400 with slow_down = polling too fast
      if (response.data?.error === 'slow_down') {
        return {
          status: 'pending',
          message: 'Polling too fast. Please wait a few more seconds before trying again.',
          userId: resolvedId,
        };
      }

      // Other errors
      const errorDesc = response.data?.error_description || response.data?.error || 'Unknown error';
      console.error(`[UserAuth] Auth poll error for ${resolvedId}: ${errorDesc}`);
      return {
        status: 'error',
        message: `Authentication error: ${errorDesc}`,
        userId: resolvedId,
      };
    } catch (e: any) {
      console.error(`[UserAuth] Network error during poll for ${resolvedId}: ${e.message}`);
      return {
        status: 'error',
        message: `Network error during authentication: ${e.message}`,
        userId: resolvedId,
      };
    }
  }

  // ===========================================================================
  // Public: Get Access Token (with auto-refresh)
  // ===========================================================================

  async getAccessToken(userId?: string): Promise<string | null> {
    const resolvedId = this.resolveTokenUserId(userId);
    if (!resolvedId) return null;

    const tokenData = this.tokens.get(resolvedId);
    if (!tokenData) return null;

    // Check if token needs refresh (< 5 min remaining)
    const timeRemaining = tokenData.expiresAt - Date.now();
    if (timeRemaining < TOKEN_REFRESH_BUFFER_MS) {
      console.log(`[UserAuth] Token for ${resolvedId} expiring in ${Math.floor(timeRemaining / 1000)}s — refreshing`);

      if (tokenData.refreshToken) {
        try {
          await this.refreshUserToken(resolvedId);
          const refreshed = this.tokens.get(resolvedId);
          if (refreshed) return refreshed.accessToken;
        } catch (e: any) {
          console.error(`[UserAuth] Refresh failed for ${resolvedId}: ${e.message}`);
          // Token expired and refresh failed — user must re-authenticate
          this.tokens.delete(resolvedId);
          return null;
        }
      } else {
        // No refresh token — token will expire
        if (timeRemaining <= 0) {
          console.warn(`[UserAuth] Token expired for ${resolvedId} — no refresh token`);
          this.tokens.delete(resolvedId);
          return null;
        }
      }
    }

    return tokenData.accessToken;
  }

  // ===========================================================================
  // Public: Status Check
  // ===========================================================================

  getStatus(userId?: string): AuthStatusResult {
    if (!this.isConfigured()) {
      return {
        authenticated: false,
        message: 'Delegated auth not configured. Set PA_USER_CLIENT_ID and PA_USER_TENANT_ID.',
      };
    }

    const resolvedId = this.resolveTokenUserId(userId);

    if (!resolvedId) {
      // Check if there's a pending auth
      if (this.pendingAuths.size > 0) {
        const pending = Array.from(this.pendingAuths.values())[0];
        return {
          authenticated: false,
          message: `Authentication pending for ${pending.userId}. Visit ${pending.verificationUri} and enter code ${pending.userCode}.`,
          userId: pending.userId,
        };
      }

      return {
        authenticated: false,
        message: 'No users authenticated. Use pa-auth-start to begin device code login.',
      };
    }

    const tokenData = this.tokens.get(resolvedId);
    if (!tokenData) {
      return {
        authenticated: false,
        message: `No token found for ${resolvedId}.`,
        userId: resolvedId,
      };
    }

    const expiresIn = Math.floor((tokenData.expiresAt - Date.now()) / 1000);

    if (expiresIn <= 0) {
      return {
        authenticated: false,
        message: `Token expired for ${resolvedId}. Use pa-auth-start to re-authenticate.`,
        userId: resolvedId,
        expiresIn: 0,
      };
    }

    return {
      authenticated: true,
      message: `Authenticated as ${resolvedId}.`,
      userId: resolvedId,
      expiresIn,
    };
  }

  // ===========================================================================
  // Public: List Authenticated Users
  // ===========================================================================

  listAuthenticatedUsers(): string[] {
    return Array.from(this.tokens.keys()).filter((userId) => {
      const data = this.tokens.get(userId);
      return data && data.expiresAt > Date.now();
    });
  }

  // ===========================================================================
  // Private: Refresh Token
  // ===========================================================================

  private async refreshUserToken(userId: string): Promise<void> {
    const tokenData = this.tokens.get(userId);
    if (!tokenData?.refreshToken) {
      throw new Error(`No refresh token for ${userId}`);
    }

    console.log(`[UserAuth] Refreshing token for: ${userId}`);

    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

    const response = await axios.post(
      url,
      new URLSearchParams({
        grant_type: 'refresh_token',
        client_id: this.clientId,
        refresh_token: tokenData.refreshToken,
        scope: FLOW_SCOPE,
      }).toString(),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );

    // Azure AD returns snake_case — map directly
    const raw = response.data;

    this.tokens.set(userId, {
      accessToken: raw.access_token,
      refreshToken: raw.refresh_token || tokenData.refreshToken, // Keep old if not returned
      expiresAt: Date.now() + raw.expires_in * 1000,
      userId,
      acquiredAt: Date.now(),
    });

    console.log(`[UserAuth] ✅ Token refreshed for ${userId} — expires in ${raw.expires_in}s`);
  }

  // ===========================================================================
  // Private: User ID Resolution
  // ===========================================================================

  /** Resolve user ID from pending auths (for poll) */
  private resolvePendingUserId(userId?: string): string | null {
    if (userId) {
      const normalized = userId.toLowerCase().trim();
      return this.pendingAuths.has(normalized) ? normalized : null;
    }
    // Auto-resolve: return most recent pending
    if (this.pendingAuths.size === 1) {
      return Array.from(this.pendingAuths.keys())[0];
    }
    return null;
  }

  /** Resolve user ID from token store (for getAccessToken / getStatus) */
  private resolveTokenUserId(userId?: string): string | null {
    if (userId) {
      const normalized = userId.toLowerCase().trim();
      return this.tokens.has(normalized) ? normalized : null;
    }
    // Auto-resolve: return most recent token
    if (this.tokens.size === 1) {
      return Array.from(this.tokens.keys())[0];
    }
    // Multiple users — find most recently acquired
    if (this.tokens.size > 1) {
      let latest: string | null = null;
      let latestTime = 0;
      for (const [id, data] of this.tokens.entries()) {
        if (data.acquiredAt > latestTime) {
          latestTime = data.acquiredAt;
          latest = id;
        }
      }
      return latest;
    }
    return null;
  }
}
