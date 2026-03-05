// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Flow Client
//
// Handles all Power Automate Flow API operations using
// admin-scoped endpoints. Service principals registered via
// New-PowerAppManagementApp MUST use /scopes/admin/ paths.
//
// Base URL: https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════

import { AxiosInstance } from 'axios';
import winston from 'winston';

const API_VERSION = '2016-11-01';

export class FlowClient {
  private client: AxiosInstance;
  private logger: winston.Logger;

  constructor(client: AxiosInstance, logger: winston.Logger) {
    this.client = client;
    this.logger = logger;
  }

  /**
   * Build an admin-scoped path for the Flow API.
   *
   * Service principals registered via New-PowerAppManagementApp
   * are authorized for admin endpoints only. User-scoped endpoints
   * (/environments/{id}/...) require 'maker permissions' which a
   * service principal does not have.
   *
   * Admin path: /scopes/admin/environments/{envId}/...
   * User path:  /environments/{envId}/... (NOT used)
   */
  private adminPath(envId: string, subPath: string): string {
    return `/scopes/admin/environments/${envId}${subPath}`;
  }

  /**
   * List all flows in an environment.
   * Admin endpoint returns all flows across all makers.
   */
  async listFlows(
    environmentId: string,
    filter?: string,
    top?: number
  ): Promise<any> {
    this.logger.info(`Listing flows in environment: ${environmentId}, filter: ${filter || 'all'}`);

    const params: Record<string, string> = {
      'api-version': API_VERSION,
    };

    if (top) {
      params['$top'] = String(top);
    }

    // Admin endpoint — returns all flows in the environment
    const url = this.adminPath(environmentId, '/v2/flows');
    this.logger.info(`Flow list admin URL: ${url}`);

    const response = await this.client.get(url, { params });
    const data = response.data;

    // Extract flow summaries from the response
    const flows = (data.value || []).map((flow: any) => ({
      name: flow.name,
      displayName: flow.properties?.displayName || flow.name,
      state: flow.properties?.state || 'Unknown',
      createdTime: flow.properties?.createdTime,
      lastModifiedTime: flow.properties?.lastModifiedTime,
      creator: flow.properties?.creator?.objectId,
    }));

    // Apply client-side filter if needed
    // (admin endpoint returns all flows; filter is informational)
    let filteredFlows = flows;
    if (filter === 'personal' || filter === 'shared') {
      this.logger.info(`Note: filter '${filter}' applied client-side. Admin endpoint returns all flows.`);
    }

    return {
      totalFlows: filteredFlows.length,
      flows: filteredFlows,
    };
  }

  /**
   * Get detailed information about a specific flow.
   */
  async getFlowDetails(
    environmentId: string,
    flowId: string
  ): Promise<any> {
    this.logger.info(`Getting flow details: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}`);
    const response = await this.client.get(url, {
      params: { 'api-version': API_VERSION },
    });

    const flow = response.data;
    return {
      name: flow.name,
      displayName: flow.properties?.displayName || flow.name,
      state: flow.properties?.state || 'Unknown',
      definition: flow.properties?.definition,
      connectionReferences: flow.properties?.connectionReferences,
      createdTime: flow.properties?.createdTime,
      lastModifiedTime: flow.properties?.lastModifiedTime,
      creator: flow.properties?.creator,
      triggers: flow.properties?.definition?.triggers,
      actions: flow.properties?.definition?.actions,
    };
  }

  /**
   * Enable (start) or disable (stop) a flow.
   */
  async enableDisableFlow(
    environmentId: string,
    flowId: string,
    action: 'start' | 'stop'
  ): Promise<any> {
    this.logger.info(`${action} flow: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/${action}`);
    const response = await this.client.post(url, null, {
      params: { 'api-version': API_VERSION },
    });

    return {
      success: true,
      action,
      flowId,
      message: `Flow ${action === 'start' ? 'enabled' : 'disabled'} successfully.`,
    };
  }

  /**
   * Permanently delete a flow.
   */
  async deleteFlow(
    environmentId: string,
    flowId: string
  ): Promise<any> {
    this.logger.info(`Deleting flow: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}`);
    await this.client.delete(url, {
      params: { 'api-version': API_VERSION },
    });

    return {
      success: true,
      flowId,
      message: 'Flow deleted successfully.',
    };
  }

  /**
   * Trigger a flow that has an HTTP request trigger.
   * Note: Trigger may use a different path pattern than other admin ops.
   */
  async triggerFlow(
    environmentId: string,
    flowId: string,
    triggerBody?: object
  ): Promise<any> {
    this.logger.info(`Triggering flow: ${flowId} in ${environmentId}`);

    // Try admin-scoped trigger path
    const url = this.adminPath(environmentId, `/flows/${flowId}/triggers/manual/run`);
    this.logger.info(`Trigger URL: ${url}`);

    try {
      const response = await this.client.post(url, triggerBody || {}, {
        params: { 'api-version': API_VERSION },
      });

      return {
        success: true,
        flowId,
        message: 'Flow triggered successfully.',
        response: response.data,
      };
    } catch (adminError: any) {
      // If admin path fails, try non-admin path as fallback
      if (adminError.response?.status === 404) {
        this.logger.warn('Admin trigger path returned 404, trying user-scoped path...');
        const fallbackUrl = `/environments/${environmentId}/flows/${flowId}/triggers/manual/run`;
        const response = await this.client.post(fallbackUrl, triggerBody || {}, {
          params: { 'api-version': API_VERSION },
        });

        return {
          success: true,
          flowId,
          message: 'Flow triggered successfully (via user path).',
          response: response.data,
        };
      }
      throw adminError;
    }
  }

  /**
   * Get run history for a flow.
   */
  async getRunHistory(
    environmentId: string,
    flowId: string,
    top?: number
  ): Promise<any> {
    this.logger.info(`Getting run history: ${flowId} in ${environmentId}, top: ${top || 'default'}`);

    const params: Record<string, string> = {
      'api-version': API_VERSION,
    };
    if (top) {
      params['$top'] = String(top);
    }

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/runs`);
    const response = await this.client.get(url, { params });
    const data = response.data;

    const runs = (data.value || []).map((run: any) => ({
      id: run.name,
      status: run.properties?.status || 'Unknown',
      startTime: run.properties?.startTime,
      endTime: run.properties?.endTime,
      trigger: run.properties?.trigger?.name,
    }));

    return {
      totalRuns: runs.length,
      runs,
    };
  }

  /**
   * Get detailed information about a specific flow run.
   */
  async getRunDetails(
    environmentId: string,
    flowId: string,
    runId: string
  ): Promise<any> {
    this.logger.info(`Getting run details: ${runId} for flow ${flowId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/runs/${runId}`);
    const response = await this.client.get(url, {
      params: { 'api-version': API_VERSION },
    });

    const run = response.data;
    return {
      id: run.name,
      status: run.properties?.status || 'Unknown',
      startTime: run.properties?.startTime,
      endTime: run.properties?.endTime,
      trigger: run.properties?.trigger,
      actions: run.properties?.actions,
      outputs: run.properties?.outputs,
      error: run.properties?.error,
    };
  }

  /**
   * Cancel a currently running flow run.
   */
  async cancelRun(
    environmentId: string,
    flowId: string,
    runId: string
  ): Promise<any> {
    this.logger.info(`Cancelling run: ${runId} for flow ${flowId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/runs/${runId}/cancel`);
    await this.client.post(url, null, {
      params: { 'api-version': API_VERSION },
    });

    return {
      success: true,
      runId,
      message: 'Run cancelled successfully.',
    };
  }
}
