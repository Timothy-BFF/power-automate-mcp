// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Flow Client
//
// All operations use admin-scoped endpoints:
//   /scopes/admin/environments/{envId}/...
//
// Service principals registered via New-PowerAppManagementApp
// MUST use admin paths. User-scoped paths require 'maker
// permissions' which a service principal does not have.
//
// Method signatures use flexible parameter types (any) to
// maintain backward compatibility with existing tool handlers
// that may pass options objects or positional arguments.
// Runtime type detection resolves the actual values.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════

import { AxiosInstance } from 'axios';
import winston from 'winston';

const API_VERSION = '2016-11-01';

// ── Return type interfaces ──────────────────────────────────
// Typed arrays ensure .map() callbacks get FlowSummary/RunSummary
// instead of implicit 'any', satisfying noImplicitAny.
// Index signatures allow flexible property access.

export interface FlowSummary {
  name: string;
  displayName: string;
  state: string;
  createdTime?: string;
  lastModifiedTime?: string;
  creator?: string;
  [key: string]: any;
}

export interface ListFlowsResult {
  totalFlows: number;
  flows: FlowSummary[];
}

export interface RunSummary {
  id: string;
  status: string;
  startTime?: string;
  endTime?: string;
  trigger?: string;
  [key: string]: any;
}

export interface RunHistoryResult {
  totalRuns: number;
  runs: RunSummary[];
}

export class FlowClient {
  private client: AxiosInstance;
  private logger: winston.Logger;

  constructor(client: AxiosInstance, logger: winston.Logger) {
    this.client = client;
    this.logger = logger;
  }

  /**
   * Build an admin-scoped path for the Flow API.
   * Admin path: /scopes/admin/environments/{envId}/...
   */
  private adminPath(envId: string, subPath: string): string {
    return `/scopes/admin/environments/${envId}${subPath}`;
  }

  // ── listFlows ───────────────────────────────────────────────
  // Accepts:
  //   listFlows(envId, filter?, top?)
  //   listFlows(envId, { filter?, top? })
  // ────────────────────────────────────────────────────────────
  async listFlows(
    environmentId: string,
    filterOrOptions?: any,
    top?: any
  ): Promise<ListFlowsResult> {
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

    const flows: FlowSummary[] = (data.value || []).map((flow: any) => ({
      name: flow.name,
      displayName: flow.properties?.displayName || flow.name,
      state: flow.properties?.state || 'Unknown',
      createdTime: flow.properties?.createdTime,
      lastModifiedTime: flow.properties?.lastModifiedTime,
      creator: flow.properties?.creator?.objectId,
    }));

    return { totalFlows: flows.length, flows };
  }

  // ── getFlowDetails ──────────────────────────────────────────
  // Accepts:
  //   getFlowDetails(envId, flowId)
  //   getFlowDetails(envId, { flowId? })
  //   getFlowDetails(envId, {})          ← empty options fallback
  // ────────────────────────────────────────────────────────────
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
  // Accepts:
  //   enableDisableFlow(envId, flowId, 'start'|'stop')
  //   enableDisableFlow(envId, flowId, { action: 'start'|'stop' })
  // ────────────────────────────────────────────────────────────
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
  // Accepts:
  //   triggerFlow(envId, flowId, body?)
  //   triggerFlow(envId, body)              ← body at position 2
  //   triggerFlow(envId, { flowId, triggerBody? })
  // ────────────────────────────────────────────────────────────
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
        // Body passed at position 2 without flowId
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
  // The original method name used by existing tool handlers.
  // Accepts:
  //   getFlowRuns(envId, flowId, top?)
  //   getFlowRuns(envId, flowId, { top? })
  //   getFlowRuns(envId, { flowId?, top? })
  // ────────────────────────────────────────────────────────────
  async getFlowRuns(
    environmentId: string,
    flowIdOrOptions?: any,
    topOrOptions?: any
  ): Promise<RunHistoryResult> {
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

    const runs: RunSummary[] = (data.value || []).map((run: any) => ({
      id: run.name,
      status: run.properties?.status || 'Unknown',
      startTime: run.properties?.startTime,
      endTime: run.properties?.endTime,
      trigger: run.properties?.trigger?.name,
    }));

    return { totalRuns: runs.length, runs };
  }

  // ── getRunHistory (alias for getFlowRuns) ───────────────────
  async getRunHistory(
    environmentId: string,
    flowIdOrOptions?: any,
    topOrOptions?: any
  ): Promise<RunHistoryResult> {
    return this.getFlowRuns(environmentId, flowIdOrOptions, topOrOptions);
  }

  // ── getRunDetails ───────────────────────────────────────────
  // Accepts:
  //   getRunDetails(envId, flowId, runId)
  //   getRunDetails(envId, flowId, { runId })
  //   getRunDetails(envId, { flowId, runId })
  // ────────────────────────────────────────────────────────────
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
