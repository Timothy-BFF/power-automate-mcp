import axios from 'axios';
import { TokenInfo } from '../types.js';

const REFRESH_BUFFER_MS = 5 * 60 * 1000;

export class AzureTokenManager {
  private tokens: Map<string, TokenInfo> = new Map();
  private refreshPromises: Map<string, Promise<string>> = new Map();
  private tenantId: string;
  private clientId: string;
  private clientSecret: string;

  constructor() {
    this.tenantId = (process.env.AZURE_TENANT_ID || '').trim();
    this.clientId = (process.env.AZURE_CLIENT_ID || '').trim();
    this.clientSecret = (process.env.AZURE_CLIENT_SECRET || '').trim();
    if (!this.tenantId || !this.clientId || !this.clientSecret) {
      console.warn('[TokenManager] Azure AD credentials not fully configured');
    }
  }

  async getToken(scope: string): Promise<string> {
    const cached = this.tokens.get(scope);
    if (cached && cached.expiresAt > Date.now() + REFRESH_BUFFER_MS) {
      return cached.accessToken;
    }
    const existing = this.refreshPromises.get(scope);
    if (existing) return existing;
    const refreshPromise = this.acquireToken(scope);
    this.refreshPromises.set(scope, refreshPromise);
    try {
      return await refreshPromise;
    } finally {
      this.refreshPromises.delete(scope);
    }
  }

  invalidate(scope: string): void {
    this.tokens.delete(scope);
  }

  private async acquireToken(scope: string): Promise<string> {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: scope,
    });
    try {
      const response = await axios.post(url, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });
      const { access_token, expires_in } = response.data;
      this.tokens.set(scope, {
        accessToken: access_token,
        expiresAt: Date.now() + (expires_in * 1000),
        scope,
      });
      console.log(`[TokenManager] Token acquired for scope: ${scope}, TTL: ${expires_in}s`);
      return access_token;
    } catch (error: any) {
      const errMsg = error.response?.data?.error_description || error.message;
      console.error(`[TokenManager] Failed to acquire token for ${scope}: ${errMsg}`);
      throw new Error(`Token acquisition failed: ${errMsg}`);
    }
  }
}
