import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';
import { UserAuthManager } from '../auth/user-auth-manager.js';

// =========================================================================
// API Configuration Constants
// =========================================================================
const BAP_SCOPE = 'https://api.bap.microsoft.com/.default';
const FLOW_SCOPE = 'https://service.flow.microsoft.com/.default';
const POWERAPPS_SCOPE = 'https://service.powerapps.com/.default';

const BAP_BASE = 'https://api.bap.microsoft.com';
const FLOW_BASE = 'https://api.flow.microsoft.com';
const POWERAPPS_BASE = 'https://api.powerapps.com';

const BAP_API_VER = '2023-06-01';
const FLOW_API_VER = '2016-11-01';
const POWERAPPS_API_VER = '2016-11-01';

const WORKFLOW_SCHEMA = 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#';

// Propagation delay settings (Dataverse write -> Flow Management API visibility)
const PROPAGATION_DELAY_MS = 5000;
const PROPAGATION_RETRY_DELAY_MS = 5000;
const PROPAGATION_MAX_RETRIES = 3;

const DATAVERSE_HEADERS: Record<string, string> = {
  'Content-Type': 'application/json; charset=utf-8',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  'Prefer': 'return=representation',
};

// =========================================================================
// Utility
// =========================================================================
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Ensures the workflow definition has all required connector parameters.
 * v3.0.2: Auto-injects both $connections and $authentication.
 */
function ensureRequiredParameters(definition: any, displayName: string): any {
  const def = { ...definition };

  if (!def.parameters) {
    def.parameters = {};
  }

  if (!def.parameters['$connections']) {
    def.parameters['$connections'] = {
      defaultValue: {},
      type: 'Object',
    };
    console.log(`[Flow] Auto-injected parameters.$connections for '${displayName}'`);
  }

  if (!def.parameters['$authentication']) {
    def.parameters['$authentication'] = {
      defaultValue: {},
      type: 'SecureObject',
    };
    console.log(`[Flow] Auto-injected parameters.$authentication for '${displayName}'`);
  }

  return def;
}

/**
 * Unwraps a definition if the agent sent it as an array or as a
 * numeric-keyed object (MCP transport artifact).
 *
 * The MCP SDK / JSON transport layer sometimes converts:
 *   [{...}]  (array)  →  {"0": {...}}  (object with numeric keys)
 * before it reaches our code. Both cause Microsoft to return:
 *   400 InvalidRequestContent: "Could not find member '0' on FlowTemplate"
 *
 * v3.0.2: Handles both real arrays AND numeric-key objects.
 */
function unwrapDefinitionIfArray(definition: any, label: string): any {
  // Case 1: Real array — [{...}]
  if (Array.isArray(definition) && definition.length === 1) {
    console.log(`[Flow] Auto-unwrapped definition array for '${label}'`);
    return definition[0];
  }

  // Case 2: MCP transport artifact — {"0": {...}}
  // The JSON transport converts [{...}] into {"0": {...}} which is an
  // object with a single numeric key "0" and no standard workflow keys.
  if (definition && typeof definition === 'object' && !Array.isArray(definition)) {
    const keys = Object.keys(definition);
    const hasNumericZero = keys.includes('0');
    const hasWorkflowKeys = keys.some(k =>
      ['triggers', 'actions', '$schema', 'contentVersion', 'parameters'].includes(k)
    );

    if (hasNumericZero && !hasWorkflowKeys) {
      const inner = definition['0'];
      if (inner && typeof inner === 'object') {
        console.log(`[Flow] Auto-unwrapped numeric-key definition for '${label}' (MCP transport artifact: {"0": {...}})`);
        return inner;
      }
    }
  }

  return definition;
}

export class PowerPlatformClient {
  private tm: AzureTokenManager;
  private userAuth: UserAuthManager | null;
  private dataverseUrlCache: Map<string, string | null> = new Map();

  constructor(tokenManager: AzureTokenManager, userAuthManager?: UserAuthManager) {
    this.tm = tokenManager;
    this.userAuth = userAuthManager || null;

    if (this.userAuth) {
      console.log('[PowerPlatformClient] Dual-token mode: service principal (read) + delegated user (write)');
    }
  }

  setUserAuthManager(userAuth: UserAuthManager): void {
    this.userAuth = userAuth;
    console.log('[PowerPlatformClient] User auth manager connected for delegated write operations');
  }

  // =========================================================================
  // Private Transport Methods — Service Principal (Read Operations)
  // =========================================================================

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

  private async dataverseRequest(envId: string, path: string, method: string = 'GET', data?: any): Promise<any> {
    const instanceUrl = await this.discoverDataverseUrl(envId);
    if (!instanceUrl) {
      throw new Error(
        `Environment '${envId}' does not have a linked Dataverse instance. ` +
        `Flow creation/update via service principal requires Dataverse.`
      );
    }
    const scope = `${instanceUrl}/.default`;
    const url = `${instanceUrl}${path}`;
    console.log(`[Dataverse] ${method} ${url}`);
    return this.request(url, method, scope, data, DATAVERSE_HEADERS);
  }

  private async request(
    url: string,
    method: string,
    scope: string,
    data?: any,
    extraHeaders?: Record<string, string>
  ): Promise<any> {
    const token = await this.tm.getToken(scope);
    const headers: Record<string, string> = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(extraHeaders || {}),
    };

    try {
      const resp = await axios({ method, url, data, headers });
      if (resp.status === 204) {
        const entityId = resp.headers?.['odata-entityid'] || '';
        const guidMatch = entityId.match(/\(([0-9a-f-]+)\)/i);
        return guidMatch
          ? { workflowid: guidMatch[1], _status: 204 }
          : { _status: 204, success: true };
      }
      return resp.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        console.log(`[API] 401 received for ${scope}, refreshing token...`);
        this.tm.invalidate(scope);
        const freshToken = await this.tm.getToken(scope);
        headers['Authorization'] = `Bearer ${freshToken}`;
        const retry = await axios({ method, url, data, headers });
        if (retry.status === 204) {
          const entityId = retry.headers?.['odata-entityid'] || '';
          const guidMatch = entityId.match(/\(([0-9a-f-]+)\)/i);
          return guidMatch
            ? { workflowid: guidMatch[1], _status: 204 }
            : { _status: 204, success: true };
        }
        return retry.data;
      }
      if (error.response) {
        console.error(`[API] Error ${error.response.status} on ${method} ${url}`);
        console.error(`[API] Response body:`, JSON.stringify(error.response.data, null, 2));
      } else {
        console.error(`[API] Network error on ${method} ${url}:`, error.message);
      }
      throw this.fmtErr(error);
    }
  }

  private fmtErr(error: any): Error {
    if (error.response) {
      const status = error.response.status;
      const data = error.response.data;
      const msg = data?.error?.message || data?.error?.code || data?.message || JSON.stringify(data);
      return new Error(`API Error ${status}: ${msg}`);
    }
    return new Error(`Request failed: ${error.message}`);
  }

  // =========================================================================
  // Private Transport — Delegated User Token (Write Operations)
  // =========================================================================

  private async userFlowRequest(
    path: string,
    method: string = 'GET',
    data?: any,
    userId?: string
  ): Promise<any> {
    const url = path.startsWith('http') ? path : `${FLOW_BASE}${path}`;

    if (this.userAuth) {
      const userToken = await this.userAuth.getAccessToken(userId);
      if (userToken) {
        const resolvedUser = userId || this.userAuth.getDefaultUserId() || 'delegated-user';
        console.log(`[Flow] ${method} ${url} [delegated: ${resolvedUser}]`);

        const headers: Record<string, string> = {
          'Authorization': `Bearer ${userToken}`,
          'Content-Type': 'application/json',
        };

        try {
          const resp = await axios({ method, url, data, headers });
          console.log(`[Flow] \u2705 ${method} ${url} \u2192 ${resp.status} [delegated: ${resolvedUser}]`);
          if (resp.status === 204) return { _status: 204, success: true };
          return resp.data;
        } catch (error: any) {
          if (error.response) {
            console.error(`[Flow] \u274c Error ${error.response.status} on ${method} ${url} [delegated: ${resolvedUser}]`);
            console.error(`[Flow] Response body:`, JSON.stringify(error.response.data, null, 2));
          } else {
            console.error(`[Flow] Network error on ${method} ${url} [delegated: ${resolvedUser}]:`, error.message);
          }

          if (error.response?.status === 401) {
            console.log(`[Flow] 401 on delegated request for ${resolvedUser}, attempting refresh...`);
            const freshToken = await this.userAuth.getAccessToken(userId);
            if (freshToken && freshToken !== userToken) {
              headers['Authorization'] = `Bearer ${freshToken}`;
              try {
                const retry = await axios({ method, url, data, headers });
                console.log(`[Flow] \u2705 Retry succeeded: ${method} ${url} \u2192 ${retry.status} [delegated: ${resolvedUser}]`);
                if (retry.status === 204) return { _status: 204, success: true };
                return retry.data;
              } catch (retryError: any) {
                if (retryError.response) {
                  console.error(`[Flow] \u274c Retry also failed: ${retryError.response.status} on ${method} ${url}`);
                  console.error(`[Flow] Retry response:`, JSON.stringify(retryError.response.data, null, 2));
                }
                throw this.fmtWriteErr(retryError);
              }
            }
          }
          throw this.fmtWriteErr(error);
        }
      }

      console.log('[Flow] No delegated token available for write operation');
    }

    console.log(`[Flow] ${method} ${url} [service-principal fallback]`);
    console.log('[Flow] \u26a0\ufe0f Write operations may require user auth \u2014 use pa-auth-start if this fails');
    return this.request(url, method, FLOW_SCOPE, data);
  }

  private fmtWriteErr(error: any): Error {
    if (error.response) {
      const status = error.response.status;
      const data = error.response.data;
      const msg = data?.error?.message || data?.error?.code || data?.message || JSON.stringify(data);

      if (status === 401 || status === 403) {
        return new Error(
          `API Error ${status}: ${msg}\n\n` +
          `This write operation requires delegated user authentication.\n` +
          `Use pa-auth-start to authenticate with your Microsoft account, then try again.`
        );
      }
      return new Error(`API Error ${status}: ${msg}`);
    }
    return new Error(`Request failed: ${error.message}`);
  }

  // =========================================================================
  // Dataverse URL Discovery (cached per environment)
  // =========================================================================

  private async discoverDataverseUrl(envId: string): Promise<string | null> {
    if (this.dataverseUrlCache.has(envId)) {
      const cached = this.dataverseUrlCache.get(envId)!;
      console.log(`[Dataverse] URL cache ${cached ? 'hit' : 'miss (no Dataverse)'}: ${envId} -> ${cached || 'N/A'}`);
      return cached;
    }

    try {
      console.log(`[Dataverse] Discovering instance URL for environment: ${envId}`);
      const envData = await this.bapRequest(
        `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${envId}`
      );
      const instanceUrl = envData?.properties?.linkedEnvironmentMetadata?.instanceUrl;
      if (instanceUrl) {
        const normalized = instanceUrl.replace(/\/+$/, '');
        console.log(`[Dataverse] Discovered instance URL: ${normalized}`);
        this.dataverseUrlCache.set(envId, normalized);
        return normalized;
      } else {
        console.warn(`[Dataverse] Environment '${envId}' has no linked Dataverse instance.`);
        this.dataverseUrlCache.set(envId, null);
        return null;
      }
    } catch (error: any) {
      console.error(`[Dataverse] Discovery failed for '${envId}': ${error.message}`);
      this.dataverseUrlCache.set(envId, null);
      return null;
    }
  }

  // =========================================================================
  // Dataverse ID Resolution
  // =========================================================================

  private async resolveFlowApiId(
    envId: string,
    dataverseWorkflowId: string
  ): Promise<{ flowApiId: string; dataverseId: string; resolvedVia: string }> {
    try {
      console.log(`[IdResolver] Resolving Flow API ID for Dataverse workflowid: ${dataverseWorkflowId}`);

      const result = await this.dataverseRequest(
        envId,
        `/api/data/v9.2/workflows(${dataverseWorkflowId})?$select=workflowidunique,resourceid,name`,
        'GET'
      );

      const workflowidunique = result?.workflowidunique;
      const resourceid = result?.resourceid;

      console.log(`[IdResolver] workflowidunique: ${workflowidunique || '(not found)'}`);
      console.log(`[IdResolver] resourceid:       ${resourceid || '(not found)'}`);

      if (workflowidunique) {
        console.log(`[IdResolver] Resolved via workflowidunique: ${workflowidunique}`);
        return { flowApiId: workflowidunique, dataverseId: dataverseWorkflowId, resolvedVia: 'workflowidunique' };
      }
      if (resourceid) {
        console.log(`[IdResolver] Resolved via resourceid: ${resourceid}`);
        return { flowApiId: resourceid, dataverseId: dataverseWorkflowId, resolvedVia: 'resourceid' };
      }

      console.warn(`[IdResolver] No bridge field found \u2014 using Dataverse workflowid as fallback`);
      return { flowApiId: dataverseWorkflowId, dataverseId: dataverseWorkflowId, resolvedVia: 'dataverse-workflowid' };
    } catch (error: any) {
      console.warn(`[IdResolver] Resolution failed: ${error.message}. Using fallback.`);
      return { flowApiId: dataverseWorkflowId, dataverseId: dataverseWorkflowId, resolvedVia: 'dataverse-workflowid' };
    }
  }

  // =========================================================================
  // Environments (BAP Admin API)
  // =========================================================================

  async listEnvironments(): Promise<any> {
    return this.bapRequest(
      '/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments'
    );
  }

  // =========================================================================
  // Flows — READ Operations (Flow Admin API — /scopes/admin/ path)
  // =========================================================================

  async listFlows(envId: string, filter?: string, top?: number): Promise<any> {
    let path = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/v2/flows`;
    const params: string[] = [];
    if (filter) params.push(`$filter=${filter}`);
    if (top) params.push(`$top=${top}`);
    if (params.length) path += `?${params.join('&')}`;
    path += `${path.includes('?') ? '&' : '?'}api-version=${FLOW_API_VER}`;
    return this.flowAdminRequest(path);
  }

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    const url = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`;

    for (let attempt = 0; attempt <= PROPAGATION_MAX_RETRIES; attempt++) {
      try {
        const result = await this.flowAdminRequest(url);
        if (attempt > 0) {
          console.log(`[Flow] Flow ${flowId} found after ${attempt} retries (propagation delay)`);
        }
        return result;
      } catch (error: any) {
        const is404 = error.message?.includes('404') || error.message?.includes('Could not find flow');

        if (is404 && attempt < PROPAGATION_MAX_RETRIES) {
          const delayMs = PROPAGATION_RETRY_DELAY_MS * (attempt + 1);
          console.log(`[Flow] 404 on attempt ${attempt + 1}/${PROPAGATION_MAX_RETRIES + 1} for flow ${flowId}. ` +
            `Waiting ${delayMs}ms for propagation...`);
          await sleep(delayMs);
          continue;
        }

        throw error;
      }
    }

    throw new Error(`Flow ${flowId} not found after ${PROPAGATION_MAX_RETRIES} retries (propagation timeout)`);
  }

  // =========================================================================
  // Flows — WRITE Operations (Delegated User Token)
  // =========================================================================
  //
  // DEFINITION AUTO-FIX (v3.0.2):
  //   1. Definition array/numeric-key unwrap — [{...}] or {"0":{...}} → {...}
  //   2. $schema + contentVersion injection (v3.0.1)
  //   3. Empty definition guard (v3.0.1)
  //   4. $connections parameter injection (v3.0.2)
  //   5. $authentication parameter injection (v3.0.2)
  //
  // =========================================================================

  async createFlow(
    envId: string,
    displayName: string,
    definition: any,
    state: string = 'Stopped',
    connectionReferences?: any
  ): Promise<any> {
    // --- Array / numeric-key unwrap (v3.0.2) ---
    definition = unwrapDefinitionIfArray(definition, displayName);

    // --- Robust schema injection ---
    let fullDefinition = { ...definition };
    if (!fullDefinition['$schema']) {
      fullDefinition['$schema'] = WORKFLOW_SCHEMA;
    }
    if (!fullDefinition.contentVersion) {
      fullDefinition.contentVersion = '1.0.0.0';
    }
    if (!fullDefinition.triggers) {
      fullDefinition.triggers = {};
    }
    if (!fullDefinition.actions) {
      fullDefinition.actions = {};
    }

    // --- Required parameters injection (v3.0.2) ---
    fullDefinition = ensureRequiredParameters(fullDefinition, displayName);

    // --- Empty definition guard ---
    const triggerCount = Object.keys(fullDefinition.triggers).length;
    const actionCount = Object.keys(fullDefinition.actions).length;

    if (triggerCount === 0 && actionCount === 0) {
      console.error(`[Flow] \u274c Rejected empty definition for '${displayName}' (0 triggers, 0 actions)`);
      throw new Error(
        `Invalid flow definition: both triggers and actions are empty. ` +
        `A valid Power Automate flow requires at least one trigger (e.g., Recurrence, Request, ` +
        `OpenApiConnection) and at least one action (e.g., Compose, HTTP, OpenApiConnection). ` +
        `Please provide a complete workflow definition object.`
      );
    }

    if (triggerCount === 0) {
      console.warn(`[Flow] \u26a0\ufe0f Definition for '${displayName}' has no triggers \u2014 flow may not execute`);
    }
    if (actionCount === 0) {
      console.warn(`[Flow] \u26a0\ufe0f Definition for '${displayName}' has no actions \u2014 flow will do nothing when triggered`);
    }

    const body: any = {
      properties: {
        displayName,
        definition: fullDefinition,
        state: state === 'Started' ? 'Started' : 'Stopped',
      },
    };

    if (connectionReferences && Object.keys(connectionReferences).length > 0) {
      body.properties.connectionReferences = connectionReferences;
    }

    console.log(`[Flow] Creating flow '${displayName}' in env ${envId} via Flow Management API`);
    console.log(`[Flow] State: ${body.properties.state}, Definition: ${JSON.stringify(fullDefinition).length} chars, Triggers: ${triggerCount}, Actions: ${actionCount}`);

    const result = await this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows?api-version=${FLOW_API_VER}`,
      'POST',
      body
    );

    const flowId = result?.name || 'unknown';
    const flowDisplayName = result?.properties?.displayName || displayName;
    const flowState = result?.properties?.state || state;

    console.log(`[Flow] \u2705 Flow created: ${flowId} (${flowDisplayName})`);

    return {
      status: 'created',
      flowId,
      name: flowId,
      displayName: flowDisplayName,
      state: flowState,
      definition: fullDefinition,
      connectionReferences: connectionReferences || {},
      _source: 'flow-management-api',
      _authType: this.userAuth?.hasAuthenticatedUser() ? 'delegated' : 'service-principal',
      _note: 'Created via Flow Management API \u2014 fully registered with Flow engine',
    };
  }

  async updateFlow(
    envId: string,
    flowId: string,
    updates: {
      displayName?: string;
      definition?: any;
      state?: string;
      connectionReferences?: any;
    }
  ): Promise<any> {
    const body: any = { properties: {} };

    if (updates.displayName) {
      body.properties.displayName = updates.displayName;
    }

    if (updates.definition) {
      // --- Array / numeric-key unwrap (v3.0.2) ---
      let defInput = unwrapDefinitionIfArray(updates.definition, `flow ${flowId}`);

      let def = { ...defInput };
      if (!def['$schema']) {
        def['$schema'] = WORKFLOW_SCHEMA;
      }
      if (!def.contentVersion) {
        def.contentVersion = '1.0.0.0';
      }
      if (!def.triggers) {
        def.triggers = {};
      }
      if (!def.actions) {
        def.actions = {};
      }

      // Required parameters injection (v3.0.2)
      def = ensureRequiredParameters(def, `flow ${flowId}`);

      body.properties.definition = def;
    }

    if (updates.connectionReferences) {
      body.properties.connectionReferences = updates.connectionReferences;
    }

    if (updates.state) {
      body.properties.state = updates.state === 'Started' ? 'Started' : 'Stopped';
    }

    const updateFields = Object.keys(body.properties).join(', ');
    console.log(`[Flow] Updating flow ${flowId} in env ${envId} via Flow Management API`);
    console.log(`[Flow] Update fields: ${updateFields}`);

    const result = await this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`,
      'PATCH',
      body
    );

    const updatedDisplayName = result?.properties?.displayName || updates.displayName;
    const updatedState = result?.properties?.state || updates.state;

    console.log(`[Flow] \u2705 Flow updated: ${flowId}`);

    return {
      name: result?.name || flowId,
      properties: {
        ...(updatedDisplayName ? { displayName: updatedDisplayName } : {}),
        ...(updatedState ? { state: updatedState } : {}),
        ...(updates.definition ? { definition: updates.definition } : {}),
      },
      _source: 'flow-management-api',
      _authType: this.userAuth?.hasAuthenticatedUser() ? 'delegated' : 'service-principal',
      _raw: result,
    };
  }

  // =========================================================================
  // Flow Management — LIFECYCLE (Flow Admin API)
  // =========================================================================

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

  // =========================================================================
  // Flow Runs (Flow Admin API)
  // =========================================================================

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
  // Connections (PowerApps Admin API)
  // =========================================================================

  async listConnections(envId: string): Promise<any> {
    return this.powerAppsAdminRequest(
      `/providers/Microsoft.PowerApps/scopes/admin/environments/${envId}/connections`
    );
  }
}
