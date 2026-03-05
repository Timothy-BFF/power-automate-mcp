import axios, { AxiosInstance, AxiosRequestConfig } from 'axios';
import type { Logger } from 'winston';
import { TokenManager } from '../auth/azure-token-manager.js';

const FLOW_API = 'https://api.flow.microsoft.com';
const BAP_API = 'https://api.bap.microsoft.com';
const POWERAPPS_API = 'https://api.powerapps.com';
const FLOW_API_VERSION = '2016-11-01';
const BAP_API_VERSION = '2023-06-01';

export class PowerPlatformClient {
  private tokenManager: TokenManager;
  private logger: Logger;
  private http: AxiosInstance;

  constructor(tokenManager: TokenManager, logger: Logger) {
    this.tokenManager = tokenManager;
    this.logger = logger;
    this.http = axios.create({ timeout: 30000 });
  }

  private async flowRequest(method: string, path: string, data?: any): Promise<any> {
    const token = await this.tokenManager.getFlowToken();
    const url = `${FLOW_API}${path}`;
    const reqConfig: AxiosRequestConfig = {
      method: method as any,
      url,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'x-ms-client-scope': 'admin',
      },
      data,
    };
    try {
      this.logger.debug(`Flow API: ${method} ${url}`);
      const res = await this.http(reqConfig);
      return res.data;
    } catch (err: any) {
      if (err.response?.status === 401) {
        this.logger.warn('401 on Flow API, retrying with fresh token...');
        this.tokenManager.clearTokens();
        const freshToken = await this.tokenManager.getFlowToken();
        reqConfig.headers = { ...reqConfig.headers, Authorization: `Bearer ${freshToken}` };
        const res = await this.http(reqConfig);
        return res.data;
      }
      const status = err.response?.status;
      const msg = err.response?.data?.error?.message || err.response?.data?.error || err.message;
      this.logger.error(`Flow API error: ${status} ${url}`, { message: msg });
      throw new Error(`${status || 'Error'}: ${msg}`);
    }
  }

  private async managementRequest(method: string, baseUrl: string, path: string, data?: any): Promise<any> {
    const token = await this.tokenManager.getManagementToken();
    const url = `${baseUrl}${path}`;
    const reqConfig: AxiosRequestConfig = {
      method: method as any,
      url,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'x-ms-client-scope': 'admin',
      },
      data,
    };
    try {
      this.logger.debug(`Management API: ${method} ${url}`);
      const res = await this.http(reqConfig);
      return res.data;
    } catch (err: any) {
      if (err.response?.status === 401) {
        this.logger.warn('401 on Management API, retrying with fresh token...');
        this.tokenManager.clearTokens();
        const freshToken = await this.tokenManager.getManagementToken();
        reqConfig.headers = { ...reqConfig.headers, Authorization: `Bearer ${freshToken}` };
        const res = await this.http(reqConfig);
        return res.data;
      }
      const status = err.response?.status;
      const msg = err.response?.data?.error?.message || err.response?.data?.error || err.message;
      this.logger.error(`Management API error: ${status} ${url}`, { message: msg });
      throw new Error(`${status || 'Error'}: ${msg}`);
    }
  }

  async listEnvironments(): Promise<any> {
    return this.managementRequest('GET', BAP_API,
      `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=${BAP_API_VERSION}&$select=properties.displayName,properties.environmentSku,properties.states,properties.linkedEnvironmentMetadata`);
  }

  async listFlows(envId: string, filter?: string, top?: number): Promise<any> {
    let path = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/v2/flows?api-version=${FLOW_API_VERSION}`;
    if (top) path += `&$top=${top}`;
    return this.flowRequest('GET', path);
  }

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    return this.flowRequest('GET',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VERSION}`);
  }

  async getRunHistory(envId: string, flowId: string, top?: number): Promise<any> {
    let path = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs?api-version=${FLOW_API_VERSION}`;
    if (top) path += `&$top=${top}`;
    return this.flowRequest('GET', path);
  }

  async getRunDetails(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowRequest('GET',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}?api-version=${FLOW_API_VERSION}`);
  }

  async cancelRun(envId: string, flowId: string, runId: string): Promise<any> {
    return this.flowRequest('POST',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}/cancel?api-version=${FLOW_API_VERSION}`);
  }

  async enableFlow(envId: string, flowId: string): Promise<any> {
    return this.flowRequest('POST',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/start?api-version=${FLOW_API_VERSION}`);
  }

  async disableFlow(envId: string, flowId: string): Promise<any> {
    return this.flowRequest('POST',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}/stop?api-version=${FLOW_API_VERSION}`);
  }

  async deleteFlow(envId: string, flowId: string): Promise<any> {
    return this.flowRequest('DELETE',
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VERSION}`);
  }

  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    const flow = await this.getFlowDetails(envId, flowId);
    const triggerUrl = flow?.properties?.flowTriggerUri;
    if (!triggerUrl) {
      throw new Error('Flow does not have an HTTP trigger URL. Only flows with manual/HTTP triggers can be triggered.');
    }
    const token = await this.tokenManager.getFlowToken();
    const res = await this.http.post(triggerUrl, body || {}, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    });
    return res.data || { status: 'triggered', statusCode: res.status };
  }

  async listConnections(envId: string): Promise<any> {
    return this.managementRequest('GET', POWERAPPS_API,
      `/providers/Microsoft.PowerApps/scopes/admin/environments/${envId}/connections?api-version=${FLOW_API_VERSION}`);
  }
}
