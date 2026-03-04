// ═══════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Flow Management Client
// Wraps the Power Platform Flow Management API
// Base: https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple
// ═══════════════════════════════════════════════════════════════════

import { AxiosInstance, AxiosError } from 'axios';
import winston from 'winston';

export interface FlowSummary {
  name: string;
  id: string;
  displayName: string;
  state: string;
  createdTime: string;
  lastModifiedTime: string;
  flowTriggerUri?: string;
  definition?: Record<string, unknown>;
}

export interface FlowRun {
  name: string;
  id: string;
  status: string;
  startTime: string;
  endTime?: string;
  triggerName: string;
  error?: Record<string, unknown>;
}

export interface FlowTriggerResult {
  statusCode: number;
  headers: Record<string, string>;
  body?: unknown;
}

export class FlowClient {
  private client: AxiosInstance;
  private logger: winston.Logger;

  constructor(client: AxiosInstance, logger: winston.Logger) {
    this.client = client;
    this.logger = logger;
  }

  /**
   * List all flows in an environment.
   * Supports filtering by shared/personal and top N.
   */
  async listFlows(environmentId: string, options?: {
    filter?: 'personal' | 'shared' | 'all';
    top?: number;
  }): Promise<FlowSummary[]> {
    this.logger.info(`Listing flows in environment: ${environmentId}`);
    try {
      const params: Record<string, string> = {
        'api-version': '2016-11-01',
      };

      if (options?.top) {
        params['$top'] = String(options.top);
      }

      // Choose endpoint based on filter
      let endpoint: string;
      if (options?.filter === 'shared') {
        endpoint = `/environments/${environmentId}/flows`;
        params['$filter'] = "search('team')";
      } else {
        endpoint = `/environments/${environmentId}/flows`;
      }

      const response = await this.client.get(endpoint, { params });
      const flows = response.data.value || [];

      return flows.map((f: any) => ({
        name: f.name,
        id: f.id,
        displayName: f.properties?.displayName || f.name,
        state: f.properties?.state || 'Unknown',
        createdTime: f.properties?.createdTime || '',
        lastModifiedTime: f.properties?.lastModifiedTime || '',
        flowTriggerUri: f.properties?.flowTriggerUri,
      }));
    } catch (error) {
      this.handleError('listFlows', error);
      throw error;
    }
  }

  /**
   * Get detailed information about a specific flow, including its definition.
   */
  async getFlowDetails(environmentId: string, flowId: string): Promise<FlowSummary> {
    this.logger.info(`Getting flow details: ${flowId}`);
    try {
      const response = await this.client.get(
        `/environments/${environmentId}/flows/${flowId}`,
        { params: { 'api-version': '2016-11-01' } }
      );

      const f = response.data;
      return {
        name: f.name,
        id: f.id,
        displayName: f.properties?.displayName || f.name,
        state: f.properties?.state || 'Unknown',
        createdTime: f.properties?.createdTime || '',
        lastModifiedTime: f.properties?.lastModifiedTime || '',
        flowTriggerUri: f.properties?.flowTriggerUri,
        definition: f.properties?.definition,
      };
    } catch (error) {
      this.handleError('getFlowDetails', error);
      throw error;
    }
  }

  /**
   * Enable or disable a flow.
   */
  async setFlowState(
    environmentId: string,
    flowId: string,
    state: 'Started' | 'Stopped'
  ): Promise<{ success: boolean; newState: string }> {
    this.logger.info(`Setting flow ${flowId} state to: ${state}`);
    try {
      if (state === 'Started') {
        await this.client.post(
          `/environments/${environmentId}/flows/${flowId}/start`,
          {},
          { params: { 'api-version': '2016-11-01' } }
        );
      } else {
        await this.client.post(
          `/environments/${environmentId}/flows/${flowId}/stop`,
          {},
          { params: { 'api-version': '2016-11-01' } }
        );
      }
      return { success: true, newState: state };
    } catch (error) {
      this.handleError('setFlowState', error);
      throw error;
    }
  }

  /**
   * Delete a flow.
   */
  async deleteFlow(environmentId: string, flowId: string): Promise<{ success: boolean }> {
    this.logger.info(`Deleting flow: ${flowId}`);
    try {
      await this.client.delete(
        `/environments/${environmentId}/flows/${flowId}`,
        { params: { 'api-version': '2016-11-01' } }
      );
      return { success: true };
    } catch (error) {
      this.handleError('deleteFlow', error);
      throw error;
    }
  }

  /**
   * Get run history for a flow.
   */
  async getFlowRuns(
    environmentId: string,
    flowId: string,
    options?: { top?: number; status?: string }
  ): Promise<FlowRun[]> {
    this.logger.info(`Getting run history for flow: ${flowId}`);
    try {
      const params: Record<string, string> = {
        'api-version': '2016-11-01',
      };
      if (options?.top) params['$top'] = String(options.top);
      if (options?.status) params['$filter'] = `status eq '${options.status}'`;

      const response = await this.client.get(
        `/environments/${environmentId}/flows/${flowId}/runs`,
        { params }
      );

      const runs = response.data.value || [];
      return runs.map((r: any) => ({
        name: r.name,
        id: r.id,
        status: r.properties?.status || 'Unknown',
        startTime: r.properties?.startTime || '',
        endTime: r.properties?.endTime,
        triggerName: r.properties?.trigger?.name || '',
        error: r.properties?.error,
      }));
    } catch (error) {
      this.handleError('getFlowRuns', error);
      throw error;
    }
  }

  /**
   * Get details of a specific flow run.
   */
  async getFlowRunDetails(
    environmentId: string,
    flowId: string,
    runId: string
  ): Promise<FlowRun> {
    this.logger.info(`Getting run details: ${runId} for flow: ${flowId}`);
    try {
      const response = await this.client.get(
        `/environments/${environmentId}/flows/${flowId}/runs/${runId}`,
        { params: { 'api-version': '2016-11-01' } }
      );

      const r = response.data;
      return {
        name: r.name,
        id: r.id,
        status: r.properties?.status || 'Unknown',
        startTime: r.properties?.startTime || '',
        endTime: r.properties?.endTime,
        triggerName: r.properties?.trigger?.name || '',
        error: r.properties?.error,
      };
    } catch (error) {
      this.handleError('getFlowRunDetails', error);
      throw error;
    }
  }

  /**
   * Trigger a flow that has an HTTP request trigger.
   */
  async triggerFlow(
    triggerUri: string,
    body?: Record<string, unknown>
  ): Promise<FlowTriggerResult> {
    this.logger.info('Triggering flow via HTTP trigger URI');
    try {
      // Trigger URIs are fully qualified — use axios directly, not the scoped client
      const response = await this.client.post(triggerUri, body || {});
      return {
        statusCode: response.status,
        headers: response.headers as Record<string, string>,
        body: response.data,
      };
    } catch (error) {
      this.handleError('triggerFlow', error);
      throw error;
    }
  }

  /**
   * Cancel a running flow run.
   */
  async cancelFlowRun(
    environmentId: string,
    flowId: string,
    runId: string
  ): Promise<{ success: boolean }> {
    this.logger.info(`Cancelling run: ${runId} for flow: ${flowId}`);
    try {
      await this.client.post(
        `/environments/${environmentId}/flows/${flowId}/runs/${runId}/cancel`,
        {},
        { params: { 'api-version': '2016-11-01' } }
      );
      return { success: true };
    } catch (error) {
      this.handleError('cancelFlowRun', error);
      throw error;
    }
  }

  private handleError(operation: string, error: unknown): void {
    const axiosErr = error as AxiosError;
    const errData = axiosErr.response?.data as Record<string, unknown> | undefined;
    this.logger.error(`FlowClient.${operation} failed`, {
      status: axiosErr.response?.status,
      statusText: axiosErr.response?.statusText,
      error: errData?.error,
      message: errData?.message || axiosErr.message,
    });
  }
}
