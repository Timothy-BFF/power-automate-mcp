// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Azure AD Token Manager
//
// Manages OAuth2 client_credentials tokens for two scopes:
//   - 'flow'       → Power Automate Flow API
//   - 'management' → Power Platform BAP API
//
// Implements the "No-Bother Protocol":
//   - Auto-refresh when token < 5 min from expiry
//   - 401 retry interceptor with fresh token
//   - 429 rate-limit backoff
//   - Concurrent request coalescing (one refresh per scope)
//
// IMPORTANT: The Flow API requires a custom x-ms-client-scope
// header for service principal authentication. This header is
// injected in the REQUEST INTERCEPTOR (not just axios defaults)
// to guarantee it reaches the API on every request.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════

import axios, { AxiosInstance, AxiosError, InternalAxiosRequestConfig } from 'axios';
import winston from 'winston';

export type TokenScope = 'flow' | 'management';

interface TokenEntry {
  accessToken: string;
  expiresAt: number;
  refreshPromise: Promise<string> | null;
}

export interface TokenManagerConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  tokenEndpoint: string;
  scopes: Record<TokenScope, string>;
}

export class AzureTokenManager {
  private config: TokenManagerConfig;
  private tokens: Map<TokenScope, TokenEntry> = new Map();
  private logger: winston.Logger;
  private headerLoggedPerScope: Set<TokenScope> = new Set();

  // ── Static factory (called by index.ts) ──
  static initialize(config: TokenManagerConfig, logger: winston.Logger): AzureTokenManager {
    return new AzureTokenManager(config, logger);
  }

  private constructor(config: TokenManagerConfig, logger: winston.Logger) {
    this.config = config;
    this.logger = logger;
  }

  /**
   * Get a valid token for the specified scope.
   * Returns cached token if still valid (> 5 min remaining).
   * Otherwise triggers a refresh.
   */
  async getToken(scope: TokenScope): Promise<string> {
    const entry = this.tokens.get(scope);

    // If we have a valid token with > 5 min remaining, return it
    if (entry?.accessToken && entry.expiresAt > Date.now() + 5 * 60 * 1000) {
      return entry.accessToken;
    }

    // If a refresh is already in progress, wait for it
    if (entry?.refreshPromise) {
      return entry.refreshPromise;
    }

    // Initiate refresh
    return this.refreshToken(scope);
  }

  /**
   * Force a token refresh for the specified scope.
   * Used by the 401 retry interceptor.
   */
  async forceRefresh(scope: TokenScope): Promise<string> {
    this.logger.info(`Force-refreshing token for scope: ${scope}`);
    this.tokens.delete(scope);
    return this.refreshToken(scope);
  }

  /**
   * Acquire a new token from Azure AD.
   */
  private async refreshToken(scope: TokenScope): Promise<string> {
    const scopeUrl = this.config.scopes[scope];
    this.logger.info(`Acquiring Azure AD token for scope: ${scope} (${scopeUrl})`);

    const refreshPromise = this.acquireToken(scope, scopeUrl);

    // Store the promise so concurrent requests can wait
    const currentEntry = this.tokens.get(scope);
    if (currentEntry) {
      currentEntry.refreshPromise = refreshPromise;
    } else {
      this.tokens.set(scope, {
        accessToken: '',
        expiresAt: 0,
        refreshPromise,
      });
    }

    try {
      const token = await refreshPromise;
      return token;
    } catch (error) {
      // Clear failed entry
      this.tokens.delete(scope);
      throw error;
    }
  }

  private async acquireToken(scope: TokenScope, scopeUrl: string): Promise<string> {
    // Use the tokenEndpoint from config, fall back to constructing from tenantId
    const endpoint = this.config.tokenEndpoint
      ? this.config.tokenEndpoint.trim()
      : `https://login.microsoftonline.com/${this.config.tenantId.trim()}/oauth2/v2.0/token`;

    try {
      const params = new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: this.config.clientId.trim(),
        client_secret: this.config.clientSecret.trim(),
        scope: scopeUrl.trim(),
      });

      const response = await axios.post(endpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 15000,
      });

      const { access_token, expires_in } = response.data;
      const expiresAt = Date.now() + (expires_in * 1000);

      this.tokens.set(scope, {
        accessToken: access_token,
        expiresAt,
        refreshPromise: null,
      });

      const expiresInMin = Math.round(expires_in / 60);
      this.logger.info(`Token acquired for scope '${scope}', expires in ${expiresInMin} min`);

      return access_token;
    } catch (error) {
      const axiosErr = error as AxiosError;
      const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
      this.logger.error('Failed to acquire Azure AD token', {
        status: axiosErr.response?.status,
        error: errData?.error,
        description: errData?.error_description,
      });
      throw new Error(
        `Azure AD token acquisition failed: ${errData?.error_description || axiosErr.message}`
      );
    }
  }

  /**
   * Create an axios instance with automatic token injection
   * and 401 retry logic for the specified scope.
   *
   * CRITICAL: The x-ms-client-scope header is injected in the
   * REQUEST INTERCEPTOR, not just in axios.create() defaults.
   * Some axios versions do not properly merge custom headers
   * from create() options into individual request configs.
   * The interceptor guarantees the header is on every request.
   */
  createAuthenticatedClient(scope: TokenScope, baseURL: string): AxiosInstance {
    const client = axios.create({
      baseURL,
      timeout: 30000,
      headers: {
        'Content-Type': 'application/json',
      },
    });

    // Store scope URL for use in interceptors
    const scopeUrl = this.config.scopes[scope];

    if (scope === 'flow') {
      this.logger.info(`Flow client will include x-ms-client-scope: ${scopeUrl}`);
    }

    // ── Request interceptor: inject Bearer token + x-ms-client-scope ──
    client.interceptors.request.use(async (config: InternalAxiosRequestConfig) => {
      // Inject Bearer token
      const token = await this.getToken(scope);
      config.headers.Authorization = `Bearer ${token}`;

      // CRITICAL: Inject x-ms-client-scope for Flow API
      // This header MUST be present on every Flow API request when
      // using service principal (client_credentials) authentication.
      // Without it, the API returns 401 ClientScopeAuthorizationFailed:
      // "The x-ms-client-scope header must not be null or empty."
      if (scope === 'flow') {
        config.headers['x-ms-client-scope'] = scopeUrl;
      }

      // Diagnostic: log outgoing headers once per scope
      if (!this.headerLoggedPerScope.has(scope)) {
        this.headerLoggedPerScope.add(scope);
        const headerKeys = Object.keys(config.headers.toJSON ? config.headers.toJSON() : config.headers);
        this.logger.info(`[${scope}] Outgoing request headers: ${headerKeys.join(', ')}`);
        this.logger.info(`[${scope}] x-ms-client-scope value: ${config.headers['x-ms-client-scope'] || 'NOT SET'}`);
        this.logger.info(`[${scope}] Authorization: Bearer ${token.substring(0, 20)}...`);
        this.logger.info(`[${scope}] Target: ${config.baseURL}${config.url}`);
      }

      return config;
    });

    // ── Response interceptor: 401 retry with fresh token ──
    client.interceptors.response.use(
      (response) => response,
      async (error: AxiosError) => {
        const originalRequest = error.config as InternalAxiosRequestConfig & { _retry?: boolean };

        if (error.response?.status === 401 && !originalRequest._retry) {
          originalRequest._retry = true;
          this.logger.warn(`401 received for scope '${scope}', attempting token refresh...`);

          // Log the 401 error body for diagnostics
          const errBody = error.response?.data;
          this.logger.warn(`401 error body: ${JSON.stringify(errBody)}`);

          try {
            const freshToken = await this.forceRefresh(scope);
            originalRequest.headers.Authorization = `Bearer ${freshToken}`;

            // Re-inject x-ms-client-scope on retry
            if (scope === 'flow') {
              originalRequest.headers['x-ms-client-scope'] = scopeUrl;
              this.logger.info(`[retry] x-ms-client-scope re-injected: ${scopeUrl}`);
            }

            return client.request(originalRequest);
          } catch (refreshError) {
            this.logger.error('Token refresh failed during 401 retry', { refreshError });
            throw refreshError;
          }
        }

        // Rate limiting: 429 with exponential backoff
        if (error.response?.status === 429) {
          const retryAfter = parseInt(
            error.response.headers['retry-after'] as string || '5',
            10
          );
          this.logger.warn(`Rate limited (429). Retrying after ${retryAfter}s...`);
          await new Promise((resolve) => setTimeout(resolve, retryAfter * 1000));
          return client.request(originalRequest);
        }

        throw error;
      }
    );

    return client;
  }

  /**
   * Check if the token manager can successfully acquire tokens.
   * Used for health checks.
   */
  async healthCheck(): Promise<{ flow: boolean; management: boolean }> {
    const results = { flow: false, management: false };
    try {
      await this.getToken('flow');
      results.flow = true;
    } catch { /* silent */ }
    try {
      await this.getToken('management');
      results.management = true;
    } catch { /* silent */ }
    return results;
  }
}
