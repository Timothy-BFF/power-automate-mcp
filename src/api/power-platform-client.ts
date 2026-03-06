import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';

const MGMT_SCOPE = 'https://service.powerapps.com/.default';
const FLOW_SCOPE = 'https://service.flow.microsoft.com/.default';
const API_VER = '2016-11-01';

export class PowerPlatformClient {
  private tm: AzureTokenManager;

  constructor(tokenManager: AzureTokenManager) {
    this.tm = tokenManager;
  }

  private async mgmtRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const sep = path.includes('?') ? '&' : '?';
    const url = `https://api.powerapps.com${path}${sep}api-version=${API_VER}`;
    return this.request(url, method, MGMT_SCOPE, data);
  }

  private async flowRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const url = path.startsWith('http') ? path : `https://api.flow.microsoft.com${path}`;
    return this.request(url, method, FLOW_SCOPE, data);
  }

  private async request(url: string, method: string, scope: string, data?: any): Promise<any> {
    const token = await this.tm.getToken(scope);
    try {
      const resp = await axios({ method, url, data, headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } });
      return resp.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        this.tm.invalidate(scope);
        const freshToken = await this.tm.getToken(scope);
        const retry = await axios({ method, url, data, headers: { Authorization: `Bearer ${freshToken}`, 'Content-Type': 'application/json' } });
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

  async listEnvironments(): Promise<any> {
    return this.mgmtRequest('/providers/Microsoft.PowerApps/environments');
  }

  async listFlows(envId: string, filter?: string, top?: number): Promise<any> {
    let p = `/providers/Microsoft.ProcessSimple/environments/${envId}/flows?api-version=${API_VER}`;
    if (filter && filter !== 'all') {
      const m: Record<string, string> = { personal: "search('personal')", shared: "search('team')" };
      if (m[filter]) p += `&$filter=${encodeURIComponent(m[filter])}`;
    }
    if (top) p += `&$top=${top}`;
    return this.flowRequest(p);
  }

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}?api-version=${API_VER}`);
  }

  async enableDisableFlow(envId: string, flowId: string, action: 'start' | 'stop'): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/${action}?api-version=${API_VER}`, 'POST');
  }

  async deleteFlow(envId: string, flowId: string): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}?api-version=${API_VER}`, 'DELETE');
  }

  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/triggers/manual/run?api-version=${API_VER}`, 'POST', body || {});
  }

  async getRunHistory(envId: string, flowId: string, top?: number): Promise<any> {
    let p = `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/runs?api-version=${API_VER}`;
    if (top) p += `&$top=${top}`;
    return this.flowRequest(p);
  }

  async getRunDetails(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/runs/${runId}?api-version=${API_VER}`);
  }

  async cancelRun(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowRequest(`/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/runs/${runId}/cancel?api-version=${API_VER}`, 'POST');
  }

  async listConnections(envId: string): Promise<any> {
    return this.mgmtRequest(`/providers/Microsoft.PowerApps/environments/${envId}/connections`);
  }
}
