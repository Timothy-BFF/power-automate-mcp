// ═══════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Environment Client
// Wraps the Power Platform BAP API for environment management
// Base: https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform
// ═══════════════════════════════════════════════════════════════════

import { AxiosInstance, AxiosError } from 'axios';
import winston from 'winston';

export interface EnvironmentSummary {
  name: string;
  id: string;
  displayName: string;
  location: string;
  type: string;
  state: string;
  createdTime: string;
  isDefault: boolean;
}

export class EnvironmentClient {
  private client: AxiosInstance;
  private logger: winston.Logger;

  constructor(client: AxiosInstance, logger: winston.Logger) {
    this.client = client;
    this.logger = logger;
  }

  /**
   * List all Power Platform environments accessible to the service principal.
   */
  async listEnvironments(): Promise<EnvironmentSummary[]> {
    this.logger.info('Listing Power Platform environments');
    try {
      const response = await this.client.get(
        '/scopes/admin/environments',
        { params: { 'api-version': '2016-11-01' } }
      );

      const environments = response.data.value || [];
      return environments.map((e: any) => ({
        name: e.name,
        id: e.id,
        displayName: e.properties?.displayName || e.name,
        location: e.location || '',
        type: e.properties?.environmentSku || 'Unknown',
        state: e.properties?.states?.management?.id || 'Unknown',
        createdTime: e.properties?.createdTime || '',
        isDefault: e.properties?.isDefault || false,
      }));
    } catch (error) {
      this.handleError('listEnvironments', error);
      throw error;
    }
  }

  /**
   * Get details of a specific environment.
   */
  async getEnvironmentDetails(environmentId: string): Promise<EnvironmentSummary> {
    this.logger.info(`Getting environment details: ${environmentId}`);
    try {
      const response = await this.client.get(
        `/scopes/admin/environments/${environmentId}`,
        { params: { 'api-version': '2016-11-01' } }
      );

      const e = response.data;
      return {
        name: e.name,
        id: e.id,
        displayName: e.properties?.displayName || e.name,
        location: e.location || '',
        type: e.properties?.environmentSku || 'Unknown',
        state: e.properties?.states?.management?.id || 'Unknown',
        createdTime: e.properties?.createdTime || '',
        isDefault: e.properties?.isDefault || false,
      };
    } catch (error) {
      this.handleError('getEnvironmentDetails', error);
      throw error;
    }
  }

  private handleError(operation: string, error: unknown): void {
    const axiosErr = error as AxiosError;
    const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
    this.logger.error(`EnvironmentClient.${operation} failed`, {
      status: axiosErr.response?.status,
      error: errData?.error,
      message: errData?.message || axiosErr.message,
    });
  }
}
