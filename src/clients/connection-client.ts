// ═══════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Connection Client
// Wraps the Power Platform API for connection management
// ═══════════════════════════════════════════════════════════════════

import { AxiosInstance, AxiosError } from 'axios';
import winston from 'winston';

export interface ConnectionSummary {
  name: string;
  id: string;
  displayName: string;
  connectorName: string;
  status: string;
  createdTime: string;
  statuses: Array<{ status: string }>;
}

export class ConnectionClient {
  private client: AxiosInstance;
  private logger: winston.Logger;

  constructor(client: AxiosInstance, logger: winston.Logger) {
    this.client = client;
    this.logger = logger;
  }

  /**
   * List all connections in an environment.
   */
  async listConnections(environmentId: string): Promise<ConnectionSummary[]> {
    this.logger.info(`Listing connections in environment: ${environmentId}`);
    try {
      const response = await this.client.get(
        `/environments/${environmentId}/connections`,
        { params: { 'api-version': '2016-11-01' } }
      );

      const connections = response.data.value || [];
      return connections.map((c: any) => ({
        name: c.name,
        id: c.id,
        displayName: c.properties?.displayName || c.name,
        connectorName: c.properties?.apiId || '',
        status: c.properties?.statuses?.[0]?.status || 'Unknown',
        createdTime: c.properties?.createdTime || '',
        statuses: c.properties?.statuses || [],
      }));
    } catch (error) {
      this.handleError('listConnections', error);
      throw error;
    }
  }

  /**
   * Get details of a specific connection.
   */
  async getConnectionDetails(
    environmentId: string,
    connectionName: string
  ): Promise<ConnectionSummary> {
    this.logger.info(`Getting connection details: ${connectionName}`);
    try {
      const response = await this.client.get(
        `/environments/${environmentId}/connections/${connectionName}`,
        { params: { 'api-version': '2016-11-01' } }
      );

      const c = response.data;
      return {
        name: c.name,
        id: c.id,
        displayName: c.properties?.displayName || c.name,
        connectorName: c.properties?.apiId || '',
        status: c.properties?.statuses?.[0]?.status || 'Unknown',
        createdTime: c.properties?.createdTime || '',
        statuses: c.properties?.statuses || [],
      };
    } catch (error) {
      this.handleError('getConnectionDetails', error);
      throw error;
    }
  }

  /**
   * Delete a connection.
   */
  async deleteConnection(
    environmentId: string,
    connectionName: string
  ): Promise<{ success: boolean }> {
    this.logger.info(`Deleting connection: ${connectionName}`);
    try {
      await this.client.delete(
        `/environments/${environmentId}/connections/${connectionName}`,
        { params: { 'api-version': '2016-11-01' } }
      );
      return { success: true };
    } catch (error) {
      this.handleError('deleteConnection', error);
      throw error;
    }
  }

  private handleError(operation: string, error: unknown): void {
    const axiosErr = error as AxiosError;
    const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
    this.logger.error(`ConnectionClient.${operation} failed`, {
      status: axiosErr.response?.status,
      error: errData?.error,
      message: errData?.message || axiosErr.message,
    });
  }
}
