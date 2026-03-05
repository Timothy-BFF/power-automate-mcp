// ═══════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Environment Client
// Wraps the Power Platform BAP API for environment management
// Base: https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform
//
// IMPORTANT: The BAP API has two endpoint tiers:
//   /scopes/admin/environments  → Requires Power Platform Admin role
//   /environments               → Lists environments the caller has access to
//
// This client tries the non-admin endpoint first (broader compatibility),
// then falls back to the admin endpoint if available.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
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
   *
   * Strategy: Try non-admin endpoint first, fall back to admin endpoint.
   * The non-admin endpoint (/environments) lists environments the caller
   * has access to without requiring the Power Platform Admin role.
   * The admin endpoint (/scopes/admin/environments) requires the service
   * principal to be registered as a Power Platform Admin.
   */
  async listEnvironments(): Promise<EnvironmentSummary[]> {
    this.logger.info('Listing Power Platform environments');

    // Strategy 1: Try non-admin endpoint first (works without admin role)
    try {
      this.logger.info('Attempting non-admin environment listing (/environments)');
      const response = await this.client.get(
        '/environments',
        { params: { 'api-version': '2016-11-01' } }
      );

      const environments = response.data.value || [];
      this.logger.info(`Non-admin endpoint returned ${environments.length} environment(s)`);
      return this.mapEnvironments(environments);
    } catch (nonAdminError) {
      const axiosErr = nonAdminError as AxiosError;
      this.logger.warn('Non-admin environment listing failed, trying admin endpoint...', {
        status: axiosErr.response?.status,
        message: axiosErr.message,
      });
    }

    // Strategy 2: Fall back to admin endpoint
    try {
      this.logger.info('Attempting admin environment listing (/scopes/admin/environments)');
      const response = await this.client.get(
        '/scopes/admin/environments',
        { params: { 'api-version': '2016-11-01' } }
      );

      const environments = response.data.value || [];
      this.logger.info(`Admin endpoint returned ${environments.length} environment(s)`);
      return this.mapEnvironments(environments);
    } catch (adminError) {
      const axiosErr = adminError as AxiosError;
      const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
      const errorDetail = errData?.error as Record<string, unknown> | undefined;

      this.logger.error('Both environment listing endpoints failed', {
        status: axiosErr.response?.status,
        error: errData?.error,
        message: errData?.message || axiosErr.message,
      });

      // Provide actionable error message
      if (axiosErr.response?.status === 403) {
        throw new Error(
          `403 Forbidden: The service principal does not have permission to list environments. ` +
          `To fix this, register the service principal as an Application User in the Power Platform Admin Center ` +
          `(admin.powerplatform.microsoft.com → Environments → [your env] → Settings → Application users). ` +
          `For admin-level access, assign the Power Platform Administrator role in Microsoft Entra. ` +
          `Original error: ${errorDetail?.message || axiosErr.message}`
        );
      }

      throw adminError;
    }
  }

  /**
   * Get details of a specific environment.
   * Tries non-admin endpoint first, falls back to admin.
   */
  async getEnvironmentDetails(environmentId: string): Promise<EnvironmentSummary> {
    this.logger.info(`Getting environment details: ${environmentId}`);

    // Try non-admin first
    try {
      const response = await this.client.get(
        `/environments/${environmentId}`,
        { params: { 'api-version': '2016-11-01' } }
      );
      return this.mapSingleEnvironment(response.data);
    } catch (nonAdminError) {
      this.logger.warn('Non-admin environment details failed, trying admin endpoint...');
    }

    // Fall back to admin
    try {
      const response = await this.client.get(
        `/scopes/admin/environments/${environmentId}`,
        { params: { 'api-version': '2016-11-01' } }
      );
      return this.mapSingleEnvironment(response.data);
    } catch (error) {
      this.handleError('getEnvironmentDetails', error);
      throw error;
    }
  }

  /**
   * Map API response array to EnvironmentSummary objects.
   */
  private mapEnvironments(environments: any[]): EnvironmentSummary[] {
    return environments.map((e: any) => this.mapSingleEnvironment(e));
  }

  /**
   * Map a single API response to EnvironmentSummary.
   */
  private mapSingleEnvironment(e: any): EnvironmentSummary {
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
  }

  private handleError(operation: string, error: unknown): void {
    const axiosErr = error as AxiosError;
    const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
    this.logger.error(`EnvironmentClient.${operation} failed`, axiosErr.message, {
      status: axiosErr.response?.status,
      error: errData?.error,
      message: errData?.message || axiosErr.message,
    });
  }
}
