// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Flow Client
//
// All operations use admin-scoped endpoints:
//   /scopes/admin/environments/{envId}/...
//
// IMPORTANT: Return types are Promise<any> and methods return
// RAW arrays (not wrapped objects) because the existing tool
// handlers call .length and .map() directly on the result.
//
// Method aliases:
//   getFlowRunDetails() → getRunDetails()  (handler compatibility)
//   getRunHistory()     → getFlowRuns()    (handler compatibility)
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
   */
  private adminPath(envId: string, subPath: string): string {
    return `/scopes/admin/environments/${envId}${subPath}`;
  }

  // ── listFlows ───────────────────────────────────────────────
  // Returns: raw array of flow objects (handler calls .length/.map)
  // Accepts:
  //   listFlows(envId, filter?, top?)
  //   listFlows(envId, { filter?, top? })
  // ────────────────────────────────────────────────────────────
  async listFlows(
    environmentId: string,
    filterOrOptions?: any,
    top?: any
  ): Promise<any> {
    let filter: string | undefined;
    let topVal: number | undefined;

    if (typeof filterOrOptions === 'object' && filterOrOptions !== null) {
      filter = filterOrOptions.filter;
      topVal = filterOrOptions.top;
    } else {
      filter = filterOrOptions;
      topVal = top;
    }

    this.logger.info(`Listing flows in environment: ${environmentId}, filter: ${filter || 'all'}`);

    const params: Record<string, string> = { 'api-version': API_VERSION };
    if (topVal) params['$top'] = String(topVal);

    const url = this.adminPath(environmentId, '/v2/flows');
    this.logger.info(`Flow list admin URL: ${url}`);

    const response = await this.client.get(url, { params });
    const data = response.data;

    // Return RAW ARRAY — handlers call result.length and result.map()
    const flows = (data.value || []).map((flow: any) => ({
      name: flow.name,
      displayName: flow.properties?.displayName || flow.name,
      state: flow.properties?.state || 'Unknown',
      createdTime: flow.properties?.createdTime,
      lastModifiedTime: flow.properties?.lastModifiedTime,
      creator: flow.properties?.creator?.objectId,
    }));

    return flows;
  }

  // ── getFlowDetails ──────────────────────────────────────────
  async getFlowDetails(
    environmentId: string,
    flowIdOrOptions?: any,
    extra?: any
  ): Promise<any> {
    let flowId: string | undefined;

    if (typeof flowIdOrOptions === 'string') {
      flowId = flowIdOrOptions;
    } else if (typeof flowIdOrOptions === 'object' && flowIdOrOptions?.flowId) {
      flowId = flowIdOrOptions.flowId;
    }

    if (!flowId) {
      throw new Error('flowId is required for getFlowDetails');
    }

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

  // ── enableDisableFlow ───────────────────────────────────────
  async enableDisableFlow(
    environmentId: string,
    flowId: string,
    actionOrOptions?: any
  ): Promise<any> {
    let action: string;

    if (typeof actionOrOptions === 'object' && actionOrOptions?.action) {
      action = actionOrOptions.action;
    } else {
      action = actionOrOptions;
    }

    if (action !== 'start' && action !== 'stop') {
      throw new Error('action must be "start" or "stop"');
    }

    this.logger.info(`${action} flow: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/${action}`);
    await this.client.post(url, null, {
      params: { 'api-version': API_VERSION },
    });

    return {
      success: true,
      action,
      flowId,
      message: `Flow ${action === 'start' ? 'enabled' : 'disabled'} successfully.`,
    };
  }

  // ── deleteFlow ──────────────────────────────────────────────
  async deleteFlow(environmentId: string, flowId: string): Promise<any> {
    this.logger.info(`Deleting flow: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}`);
    await this.client.delete(url, {
      params: { 'api-version': API_VERSION },
    });

    return { success: true, flowId, message: 'Flow deleted successfully.' };
  }

  // ── triggerFlow ─────────────────────────────────────────────
  async triggerFlow(
    environmentId: string,
    flowIdOrBody?: any,
    body?: any
  ): Promise<any> {
    let flowId: string | undefined;
    let triggerBody: object | undefined;

    if (typeof flowIdOrBody === 'string') {
      flowId = flowIdOrBody;
      triggerBody = body;
    } else if (typeof flowIdOrBody === 'object' && flowIdOrBody !== null) {
      if (flowIdOrBody.flowId) {
        flowId = flowIdOrBody.flowId;
        triggerBody = flowIdOrBody.triggerBody || undefined;
      } else {
        triggerBody = flowIdOrBody;
      }
    }

    if (!flowId) {
      throw new Error('flowId is required for triggerFlow');
    }

    this.logger.info(`Triggering flow: ${flowId} in ${environmentId}`);

    const url = this.adminPath(environmentId, `/flows/${flowId}/triggers/manual/run`);

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
      if (adminError.response?.status === 404) {
        this.logger.warn('Admin trigger path returned 404, trying user path...');
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

  // ── getFlowRuns ─────────────────────────────────────────────
  // Returns: raw array of run objects (handler calls .length/.map)
  // Accepts:
  //   getFlowRuns(envId, flowId, top?)
  //   getFlowRuns(envId, flowId, { top? })
  //   getFlowRuns(envId, { flowId?, top? })
  // ────────────────────────────────────────────────────────────
  async getFlowRuns(
    environmentId: string,
    flowIdOrOptions?: any,
    topOrOptions?: any
  ): Promise<any> {
    let flowId: string | undefined;
    let top: number | undefined;

    if (typeof flowIdOrOptions === 'object' && flowIdOrOptions !== null) {
      flowId = flowIdOrOptions.flowId;
      top = flowIdOrOptions.top;
    } else {
      flowId = flowIdOrOptions;
      if (typeof topOrOptions === 'object' && topOrOptions !== null) {
        top = topOrOptions.top;
      } else {
        top = topOrOptions;
      }
    }

    if (!flowId) {
      throw new Error('flowId is required for getFlowRuns / getRunHistory');
    }

    this.logger.info(`Getting run history: ${flowId} in ${environmentId}, top: ${top || 'default'}`);

    const params: Record<string, string> = { 'api-version': API_VERSION };
    if (top) params['$top'] = String(top);

    const url = this.adminPath(environmentId, `/v2/flows/${flowId}/runs`);
    const response = await this.client.get(url, { params });
    const data = response.data;

    // Return RAW ARRAY — handlers call result.length and result.map()
    const runs = (data.value || []).map((run: any) => ({
      id: run.name,
      status: run.properties?.status || 'Unknown',
      startTime: run.properties?.startTime,
      endTime: run.properties?.endTime,
      trigger: run.properties?.trigger?.name,
    }));

    return runs;
  }

  // ── getRunHistory (alias for getFlowRuns) ───────────────────
  async getRunHistory(
    environmentId: string,
    flowIdOrOptions?: any,
    topOrOptions?: any
  ): Promise<any> {
    return this.getFlowRuns(environmentId, flowIdOrOptions, topOrOptions);
  }

  // ── getRunDetails ───────────────────────────────────────────
  async getRunDetails(
    environmentId: string,
    flowIdOrOptions?: any,
    runIdOrOptions?: any
  ): Promise<any> {
    let flowId: string | undefined;
    let runId: string | undefined;

    if (typeof flowIdOrOptions === 'object' && flowIdOrOptions !== null) {
      flowId = flowIdOrOptions.flowId;
      runId = flowIdOrOptions.runId;
    } else {
      flowId = flowIdOrOptions;
      if (typeof runIdOrOptions === 'object' && runIdOrOptions !== null) {
        runId = runIdOrOptions.runId;
      } else {
        runId = runIdOrOptions;
      }
    }

    if (!flowId || !runId) {
      throw new Error('flowId and runId are required for getRunDetails');
    }

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

  // ── getFlowRunDetails (alias — handler uses this name) ──────
  async getFlowRunDetails(
    environmentId: string,
    flowIdOrOptions?: any,
    runIdOrOptions?: any
  ): Promise<any> {
    return this.getRunDetails(environmentId, flowIdOrOptions, runIdOrOptions);
  }

  // ── cancelRun ───────────────────────────────────────────────
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

    return { success: true, runId, message: 'Run cancelled successfully.' };
  }
}
