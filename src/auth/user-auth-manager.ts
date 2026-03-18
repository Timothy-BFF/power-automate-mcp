/**
 * UserAuthManager — Per-User Delegated Authentication via Device Code Flow
 * 
 * Ported from Power Interpreter's MSAuthManager (Python → TypeScript)
 * Source: BolthouseFreshFoods/power-interpreter/app/microsoft/auth_manager.py
 * 
 * Architecture:
 *   - Device Code Flow (OAuth 2.0) for interactive user login
 *   - In-memory token cache (Map<userId, UserToken>)
 *   - Auto-refresh with 5-minute buffer (matches Power Interpreter pattern)
 *   - User resolution: explicit param → memory → most recent
 *   - Separate Azure AD App Registration from Service Principal
 * 
 * @version 3.0.0
 */

import axios from 'axios';

// ─── Types ────────────────────────────────────────────────────────────────────

interface UserToken {
  userId: string;
  accessToken: string;
  refreshToken: string;
  expiresAt: number;       // Unix timestamp (seconds)
  acquiredAt: number;      // When token was first acquired
  scope: string;
}

interface DeviceCodeResponse {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  interval: number;
  message: string;
}

interface PendingAuth {
  userId: string;
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number;
  interval: number;
}

export interface AuthStartResult {
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  message: string;
}

export interface AuthPollResult {
  status: 'authenticated' | 'pending' | 'expired' | 'error';
  userId?: string;
  message: string;
}

export interface AuthStatusResult {
  authenticated: boolean;
  userId?: string;
  expiresIn?: number;
  message: string;
}

// ─── Custom Errors ───────────────────────────────────────────────────────────

export class AuthRequiredError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

// ─── Constants ───────────────────────────────────────────────────────────────

const REFRESH_BUFFER_SECONDS = 300; // 5 minutes — matches Power Interpreter
const TOKEN_ENDPOINT_TEMPLATE = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';
const DEVICE_CODE_ENDPOINT_TEMPLATE = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/devicecode';

// Power Automate delegated scope (different from Power Interpreter's Graph scopes)
const PA_DELEGATED_SCOPES = 'https://service.flow.microsoft.com/.default offline_access';

// ─── UserAuthManager ─────────────────────────────────────────────────────────

export class UserAuthManager {
  private tokens: Map<string, UserToken> = new Map();
  private pendingAuths: Map<string, PendingAuth> = new Map();
  private mostRecentUserId: string | null = null;

  private clientId: string;
  private tenantId: string;
  private tokenEndpoint: string;
  private deviceCodeEndpoint: string;

  constructor() {
    this.clientId = process.env.PA_USER_CLIENT_ID || '';
    this.tenantId = process.env.PA_USER_TENANT_ID || process.env.AZURE_TENANT_ID || '';

    if (!this.clientId) {
      console.warn('[UserAuthManager] PA_USER_CLIENT_ID not set — delegated auth will be unavailable');
    }
    if (!this.tenantId) {
      console.warn('[UserAuthManager] PA_USER_TENANT_ID / AZURE_TENANT_ID not set');
    }

    this.tokenEndpoint = TOKEN_ENDPOINT_TEMPLATE.replace('{tenantId}', this.tenantId);
    this.deviceCodeEndpoint = DEVICE_CODE_ENDPOINT_TEMPLATE.replace('{tenantId}', this.tenantId);

    console.log('[UserAuthManager] Initialized — Device Code Flow ready');
    console.log(`[UserAuthManager] Tenant: ${this.tenantId}`);
    console.log(`[UserAuthManager] Client: ${this.clientId ? this.clientId.substring(0, 8) + '...' : 'NOT SET'}`);
  }

  // ─── Device Code Flow: Start ──────────────────────────────────────────────

  async startAuth(userId: string): Promise<AuthStartResult> {
    if (!this.clientId || !this.tenantId) {
      throw new Error(
        'Delegated auth not configured. Set PA_USER_CLIENT_ID and PA_USER_TENANT_ID environment variables.'
      );
    }

    console.log(`[UserAuthManager] Starting device code flow for: ${userId}`);

    try {
      const response = await axios.post(
        this.deviceCodeEndpoint,
        new URLSearchParams({
          client_id: this.clientId,
          scope: PA_DELEGATED_SCOPES,
        }).toString(),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        }
      );

      const data: DeviceCodeResponse = response.data;

      // Store pending auth state
      this.pendingAuths.set(userId, {
        userId,
        deviceCode: data.deviceCode || data.device_code,
        userCode: data.userCode || data.user_code,
        verificationUri: data.verificationUri || data.verification_uri,
        expiresAt: Math.floor(Date.now() / 1000) + (data.expiresIn || data.expires_in || 900),
        interval: data.interval || 5,
      });

      const userCode = data.userCode || data.user_code;
      const verificationUri = data.verificationUri || data.verification_uri;

      console.log(`[UserAuthManager] Device code issued for ${userId}: ${userCode}`);

      return {
        userCode,
        verificationUri,
        expiresIn: data.expiresIn || data.expires_in || 900,
        message: `Please visit ${verificationUri} and enter code: ${userCode}`,
      };
    } catch (error: any) {
      const detail = error.response?.data?.error_description || error.message;
      console.error(`[UserAuthManager] Device code request failed: ${detail}`);
      throw new Error(`Failed to start device code flow: ${detail}`);
    }
  }

  // ─── Device Code Flow: Poll ───────────────────────────────────────────────

  async pollAuth(userId?: string): Promise<AuthPollResult> {
    const resolvedUserId = this.resolvePendingUser(userId);

    if (!resolvedUserId) {
      return {
        status: 'error',
        message: 'No pending authentication found. Use pa-auth-start first.',
      };
    }

    const pending = this.pendingAuths.get(resolvedUserId);
    if (!pending) {
      return {
        status: 'error',
        message: `No pending authentication for user: ${resolvedUserId}`,
      };
    }

    // Check expiry
    if (Math.floor(Date.now() / 1000) > pending.expiresAt) {
      this.pendingAuths.delete(resolvedUserId);
      return {
        status: 'expired',
        userId: resolvedUserId,
        message: 'Device code expired. Please use pa-auth-start to begin again.',
      };
    }

    try {
      const response = await axios.post(
        this.tokenEndpoint,
        new URLSearchParams({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: this.clientId,
          device_code: pending.deviceCode,
        }).toString(),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        }
      );

      // Success — user completed login
      const tokenData = response.data;
      const now = Math.floor(Date.now() / 1000);

      const userToken: UserToken = {
        userId: resolvedUserId,
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || '',
        expiresAt: now + (tokenData.expires_in || 3599),
        acquiredAt: now,
        scope: tokenData.scope || PA_DELEGATED_SCOPES,
      };

      this.tokens.set(resolvedUserId, userToken);
      this.mostRecentUserId = resolvedUserId;
      this.pendingAuths.delete(resolvedUserId);

      console.log(`[UserAuthManager] ✓ Authenticated: ${resolvedUserId}`);
      console.log(`[UserAuthManager]   Token expires at: ${new Date(userToken.expiresAt * 1000).toISOString()}`);
      console.log(`[UserAuthManager]   Refresh token: ${userToken.refreshToken ? 'present' : 'MISSING'}`);

      return {
        status: 'authenticated',
        userId: resolvedUserId,
        message: `Successfully authenticated as ${resolvedUserId}`,
      };
    } catch (error: any) {
      const errorCode = error.response?.data?.error;

      // authorization_pending = user hasn't completed login yet (normal)
      if (errorCode === 'authorization_pending') {
        return {
          status: 'pending',
          userId: resolvedUserId,
          message: `Waiting for user to complete login at ${pending.verificationUri} with code: ${pending.userCode}`,
        };
      }

      // slow_down = polling too fast
      if (errorCode === 'slow_down') {
        return {
          status: 'pending',
          userId: resolvedUserId,
          message: 'Polling too fast — please wait a moment before trying again.',
        };
      }

      // expired_token = device code expired
      if (errorCode === 'expired_token') {
        this.pendingAuths.delete(resolvedUserId);
        return {
          status: 'expired',
          userId: resolvedUserId,
          message: 'Device code expired. Please use pa-auth-start to begin again.',
        };
      }

      // Other errors
      const detail = error.response?.data?.error_description || error.message;
      console.error(`[UserAuthManager] Poll error for ${resolvedUserId}: ${detail}`);
      return {
        status: 'error',
        userId: resolvedUserId,
        message: `Authentication error: ${detail}`,
      };
    }
  }

  // ─── Status Check ─────────────────────────────────────────────────────────

  getStatus(userId?: string): AuthStatusResult {
    const resolvedUserId = this.resolveUser(userId);

    if (!resolvedUserId) {
      // Check if there are ANY authenticated users
      if (this.tokens.size === 0) {
        return {
          authenticated: false,
          message: 'No users authenticated. Use pa-auth-start to begin device login.',
        };
      }

      // List authenticated users
      const users = Array.from(this.tokens.keys());
      return {
        authenticated: true,
        message: `Authenticated users: ${users.join(', ')}`,
      };
    }

    const token = this.tokens.get(resolvedUserId);
    if (!token) {
      // Check pending
      const pending = this.pendingAuths.get(resolvedUserId);
      if (pending) {
        return {
          authenticated: false,
          userId: resolvedUserId,
          message: `Authentication pending. Visit ${pending.verificationUri} and enter: ${pending.userCode}`,
        };
      }
      return {
        authenticated: false,
        userId: resolvedUserId,
        message: `Not authenticated. Use pa-auth-start to begin device login for ${resolvedUserId}.`,
      };
    }

    const now = Math.floor(Date.now() / 1000);
    const expiresIn = token.expiresAt - now;

    if (expiresIn <= 0) {
      return {
        authenticated: false,
        userId: resolvedUserId,
        expiresIn: 0,
        message: token.refreshToken
          ? 'Token expired but refresh token available — will auto-refresh on next request.'
          : 'Token expired and no refresh token. Use pa-auth-start to re-authenticate.',
      };
    }

    return {
      authenticated: true,
      userId: resolvedUserId,
      expiresIn,
      message: `Authenticated as ${resolvedUserId}. Token valid for ${Math.floor(expiresIn / 60)} minutes.`,
    };
  }

  // ─── Token Access (called by TokenRouter) ──────────────────────────────────

  async getAccessToken(userId?: string): Promise<string | null> {
    const resolvedUserId = this.resolveUser(userId);

    if (!resolvedUserId) {
      console.log('[UserAuthManager] No user ID resolved — no delegated token available');
      return null;
    }

    const token = this.tokens.get(resolvedUserId);
    if (!token) {
      console.log(`[UserAuthManager] No token found for: ${resolvedUserId}`);
      return null;
    }

    const now = Math.floor(Date.now() / 1000);
    const timeRemaining = token.expiresAt - now;

    // Token still valid and outside refresh buffer
    if (timeRemaining > REFRESH_BUFFER_SECONDS) {
      return token.accessToken;
    }

    // Token expired or within 5-min buffer — attempt refresh
    console.log(`[UserAuthManager] Token for ${resolvedUserId} needs refresh (${timeRemaining}s remaining)`);

    if (token.refreshToken) {
      const refreshed = await this.refreshToken(resolvedUserId, token.refreshToken);
      if (refreshed) {
        const updatedToken = this.tokens.get(resolvedUserId);
        return updatedToken?.accessToken || null;
      }
    }

    // Refresh failed or no refresh token
    if (timeRemaining > 0) {
      // Token still technically valid, use it
      console.warn(`[UserAuthManager] Refresh failed but token still valid for ${timeRemaining}s`);
      return token.accessToken;
    }

    // Fully expired, no refresh possible
    console.error(`[UserAuthManager] Token fully expired for ${resolvedUserId}, refresh failed`);
    this.tokens.delete(resolvedUserId);
    return null;
  }

  // ─── Token Refresh ────────────────────────────────────────────────────────

  private async refreshToken(userId: string, refreshToken: string): Promise<boolean> {
    console.log(`[UserAuthManager] Attempting token refresh for: ${userId}`);

    try {
      const response = await axios.post(
        this.tokenEndpoint,
        new URLSearchParams({
          grant_type: 'refresh_token',
          client_id: this.clientId,
          refresh_token: refreshToken,
          scope: PA_DELEGATED_SCOPES,
        }).toString(),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        }
      );

      const tokenData = response.data;
      const now = Math.floor(Date.now() / 1000);

      this.tokens.set(userId, {
        userId,
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || refreshToken, // Some flows return new RT
        expiresAt: now + (tokenData.expires_in || 3599),
        acquiredAt: now,
        scope: tokenData.scope || PA_DELEGATED_SCOPES,
      });

      console.log(`[UserAuthManager] ✓ Token refreshed for: ${userId}`);
      return true;
    } catch (error: any) {
      const detail = error.response?.data?.error_description || error.message;
      console.error(`[UserAuthManager] ✗ Refresh failed for ${userId}: ${detail}`);
      return false;
    }
  }

  // ─── User Resolution (matches Power Interpreter pattern) ──────────────────

  private resolveUser(userId?: string): string | null {
    // Priority 1: Explicit parameter
    if (userId) return userId;

    // Priority 2: Most recently authenticated user
    if (this.mostRecentUserId && this.tokens.has(this.mostRecentUserId)) {
      return this.mostRecentUserId;
    }

    // Priority 3: Any authenticated user (if only one)
    if (this.tokens.size === 1) {
      const [onlyUser] = this.tokens.keys();
      return onlyUser;
    }

    return null;
  }

  private resolvePendingUser(userId?: string): string | null {
    if (userId) return userId;

    // If only one pending auth, use that
    if (this.pendingAuths.size === 1) {
      const [onlyUser] = this.pendingAuths.keys();
      return onlyUser;
    }

    return null;
  }

  // ─── Admin Methods ────────────────────────────────────────────────────────

  listAuthenticatedUsers(): string[] {
    return Array.from(this.tokens.keys());
  }

  clearUserToken(userId: string): boolean {
    const existed = this.tokens.delete(userId);
    this.pendingAuths.delete(userId);
    if (this.mostRecentUserId === userId) {
      this.mostRecentUserId = null;
    }
    console.log(`[UserAuthManager] Cleared token for ${userId}: ${existed ? 'removed' : 'not found'}`);
    return existed;
  }

  isConfigured(): boolean {
    return !!(this.clientId && this.tenantId);
  }
}
