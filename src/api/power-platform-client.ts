import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';

// Service principal (New-PowerAppManagementApp) requires BAP admin scope + admin-prefixed paths
const BAP_SCOPE = 'https://api.bap.microsoft.com/.default';
const FLOW_SCOPE = 'https://service.flow.microsoft.com/.default';
const POWERAPPS_SCOPE = 'https://service.powerapps.com/.default';
const BAP_BASE = 'https://api.bap.microsoft.com';
const FLOW_BASE = 'https://api.flow.microsoft.com';
const POWERAPPS_BASE = 'https://api.powerapps.com';
const BAP_API_VER = '2023-06-01';
const FLOW_API_VER = '2016-11-01';
const POWERAPPS_API_VER = '2016-11-01';

export class PowerPlatformClient {
  private tm: AzureTokenManager;

  constructor(tokenManager: AzureTokenManager) {
    this.tm = tokenManager;
  }

  private async bapRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const sep = path.includes('?') ? '&' : '?';
    const url = `${BAP_BASE}${path}${sep}api-version=${BAP_API_VER}`;
    console.log(`[BAP] ${method} ${url}`);
    return this.request(url, method, BAP_SCOPE, data);
  }

  private async flowAdminRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const url = path.startsWith('http') ? path : `${FLOW_BASE}${path}`;
    console.log(`[Flow] ${method} ${url}`);
    return this.request(url, method, FLOW_SCOPE, data);
  }

  private async powerAppsAdminRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const sep = path.includes('?') ? '&' : '?';
    const url = `${POWERAPPS_BASE}${path}${sep}api-version=${POWERAPPS_API_VER}`;
    console.log(`[PowerApps] ${method} ${url}`);
    return this.request(url, method, POWERAPPS_SCOPE, data);
  }

  private async request(url: string, method: string, scope: string, data?: any): Promise<any> {
    const token = await this.tm.getToken(scope);
    try {
      const resp = await axios({
        method, url, data,
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      });
      return resp.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        console.log(`[API] 401 received for ${scope}, refreshing token...`);
        this.tm.invalidate(scope);
        const freshToken = await this.tm.getToken(scope);
        const retry = await axios({
          method, url, data,
          headers: { Authorization: `Bearer ${freshToken}`, 'Content-Type': 'application/json' },
        });
        return retry.data;
      }
      throw this.fmtErr(error);
    }
  }

  private fmtErr(e: any): Error {
    if (e.response) {
      const d = e.response.data;
      const msg = d?.error?.message || d?.message || JSON.stringify(d);
      return new Error(`API ${e.response.status}: ${msg}`);
    }
    return e;
  }

  // =========================================================================
  // Environment Management (BAP Admin API)
  // =========================================================================
  async listEnvironments(): Promise<any> {
    return this.bapRequest('/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments');
  }

  // =========================================================================
  // Flow Management (Flow Admin API with /scopes/admin/ prefix)
  // =========================================================================
  async listFlows(envId: string, filter?: string, top?: number): Promise<any> {
    let p = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/v2/flows?api-version=${FLOW_API_VER}`;
    if (filter && filter !== 'all') {
      const m: Record<string, string> = { personal: "search('personal')", shared: "search('team')" };
      if (m[filter]) p += `&$filter=${encodeURIComponent(m[filter])}`;
    }
    if (top) p += `&$top=${top}`;
    return this.flowAdminRequest(p);
  }

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`
    );
  }

  async enableDisableFlow(envId: string, flowId: string, action: 'start' | 'stop'): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/${action}?api-version=${FLOW_API_VER}`,
      'POST'
    );
  }

  async deleteFlow(envId: string, flowId: string): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`,
      'DELETE'
    );
  }

  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/triggers/manual/run?api-version=${FLOW_API_VER}`,
      'POST',
      body || {}
    );
  }

  async getRunHistory(envId: string, flowId: string, top?: number): Promise<any> {
    let p = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs?api-version=${FLOW_API_VER}`;
    if (top) p += `&$top=${top}`;
    return this.flowAdminRequest(p);
  }

  async getRunDetails(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}?api-version=${FLOW_API_VER}`
    );
  }

  async cancelRun(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}/cancel?api-version=${FLOW_API_VER}`,
      'POST'
    );
  }

  // =========================================================================
  // Connections (PowerApps Admin API — different host + scope from BAP)
  // =========================================================================
  async listConnections(envId: string): Promise<any> {
    return this.powerAppsAdminRequest(
      `/providers/Microsoft.PowerApps/scopes/admin/environments/${envId}/connections`
    );
  }
}
