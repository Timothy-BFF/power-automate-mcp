// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Azure AD Token Manager
// "No-Bother" Protocol: Users never manually refresh tokens.
//
// Supports dual-scope tokens:
//   - Flow API scope (service.flow.microsoft.com)
//   - Management API scope (management.azure.com)
//
// Features:
//   - Proactive refresh when token has < 5 min remaining
//   - Automatic 401 retry with fresh token
//   - Thread-safe singleton per scope
// ═══════════════════════════════════════════════════════════════

import axios, { AxiosInstance, InternalAxiosRequestConfig, AxiosError } from 'axios';
import winston from 'winston';

export type TokenScope = 'flow' | 'management';

interface TokenEntry {
  accessToken: string;
  expiresAt: number; // Unix timestamp in ms
  refreshPromise: Promise<string> | null;
}

interface AzureTokenConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  tokenEndpoint: string;
  scopes: Record<TokenScope, string>;
}

export class AzureTokenManager {
  private static instance: AzureTokenManager | null = null;
  private tokens: Map<TokenScope, TokenEntry> = new Map();
  private readonly REFRESH_BUFFER_MS = 5 * 60 * 1000; // 5 minutes
  private config: AzureTokenConfig;
  private logger: winston.Logger;

  private constructor(config: AzureTokenConfig, logger: winston.Logger) {
    this.config = config;
    this.logger = logger;
  }

  static initialize(config: AzureTokenConfig, logger: winston.Logger): AzureTokenManager {
    if (!AzureTokenManager.instance) {
      AzureTokenManager.instance = new AzureTokenManager(config, logger);
    }
    return AzureTokenManager.instance;
  }

  static getInstance(): AzureTokenManager {
    if (!AzureTokenManager.instance) {
      throw new Error('AzureTokenManager not initialized. Call initialize() first.');
    }
    return AzureTokenManager.instance;
  }

  /**
   * Get a valid access token for the specified scope.
   * Automatically refreshes if expired or nearing expiry.
   */
  async getToken(scope: TokenScope): Promise<string> {
    const entry = this.tokens.get(scope);

    // If we have a valid token with buffer, return it
    if (entry && entry.expiresAt - Date.now() > this.REFRESH_BUFFER_MS) {
      return entry.accessToken;
    }

    // If a refresh is already in progress, wait for it
    if (entry?.refreshPromise) {
      this.logger.debug(`Token refresh already in progress for scope: ${scope}`);
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

    const refreshPromise = this.acquireToken(scopeUrl);

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

  private async acquireToken(scopeUrl: string): Promise<string> {
    try {
      const params = new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: this.config.clientId,
        client_secret: this.config.clientSecret,
        scope: scopeUrl,
      });

      const response = await axios.post(this.config.tokenEndpoint, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 15000,
      });

      const { access_token, expires_in } = response.data;
      const expiresAt = Date.now() + (expires_in * 1000);

      // Determine which scope this belongs to
      const scope = Object.entries(this.config.scopes).find(
        ([, url]) => url === scopeUrl
      )?.[0] as TokenScope;

      if (scope) {
        this.tokens.set(scope, {
          accessToken: access_token,
          expiresAt,
          refreshPromise: null,
        });

        const expiresInMin = Math.round(expires_in / 60);
        this.logger.info(`Token acquired for scope '${scope}', expires in ${expiresInMin} min`);
      }

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
   */
  createAuthenticatedClient(scope: TokenScope, baseURL: string): AxiosInstance {
    const client = axios.create({
      baseURL,
      timeout: 30000,
      headers: { 'Content-Type': 'application/json' },
    });

    // Request interceptor: inject Bearer token
    client.interceptors.request.use(async (config: InternalAxiosRequestConfig) => {
      const token = await this.getToken(scope);
      config.headers.Authorization = `Bearer ${token}`;
      return config;
    });

    // Response interceptor: 401 retry with fresh token
    client.interceptors.response.use(
      (response) => response,
      async (error: AxiosError) => {
        const originalRequest = error.config as InternalAxiosRequestConfig & { _retry?: boolean };

        if (error.response?.status === 401 && !originalRequest._retry) {
          originalRequest._retry = true;
          this.logger.warn(`401 received for scope '${scope}', attempting token refresh...`);

          try {
            const freshToken = await this.forceRefresh(scope);
            originalRequest.headers.Authorization = `Bearer ${freshToken}`;
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
