import axios from 'axios';
import { TokenStore } from './token-store.js';
import type { Logger } from 'winston';

const FLOW_SCOPE = 'https://service.flow.microsoft.com/.default';
const MANAGEMENT_SCOPE = 'https://service.powerapps.com/.default';

export class TokenManager {
  private store = new TokenStore();
  private tenantId: string;
  private clientId: string;
  private clientSecret: string;
  private logger: Logger;

  constructor(tenantId: string, clientId: string, clientSecret: string, logger: Logger) {
    this.tenantId = tenantId;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.logger = logger;
  }

  private async acquireToken(scope: string): Promise<string> {
    const cached = this.store.get(scope);
    if (cached) return cached;

    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope,
    });

    try {
      const response = await axios.post(url, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });
      const { access_token, expires_in } = response.data;
      this.store.set(scope, access_token, expires_in);
      this.logger.info(`Token acquired for scope: ${scope}, TTL: ${expires_in}s`);
      return access_token;
    } catch (err: any) {
      this.logger.error(`Token acquisition failed for scope: ${scope}`, {
        status: err.response?.status,
        error: err.response?.data?.error,
        description: err.response?.data?.error_description,
      });
      throw new Error(`Failed to acquire token for ${scope}: ${err.response?.data?.error_description || err.message}`);
    }
  }

  async getFlowToken(): Promise<string> {
    return this.acquireToken(FLOW_SCOPE);
  }

  async getManagementToken(): Promise<string> {
    return this.acquireToken(MANAGEMENT_SCOPE);
  }

  clearTokens(): void {
    this.store.clear();
    this.logger.info('All cached tokens cleared');
  }
}
