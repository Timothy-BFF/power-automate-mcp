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

// Propagation delay settings (Dataverse write \u2192 Flow Management API visibility)
// IT team confirmed: 5-30 seconds is typical. We use conservative defaults.
const PROPAGATION_DELAY_MS = 5000;        // Initial wait after creation
const PROPAGATION_RETRY_DELAY_MS = 5000;  // Delay between retries
const PROPAGATION_MAX_RETRIES = 3;        // Max retries on 404 in getFlowDetails

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

  /**
   * Connect a UserAuthManager after construction.
   * Enables delegated user tokens for write operations (createFlow, updateFlow).
   */
  setUserAuthManager(userAuth: UserAuthManager): void {
    this.userAuth = userAuth;
    console.log('[PowerPlatformClient] User auth manager connected for delegated write operations');
  }

  // =========================================================================
  // Private Transport Methods \u2014 Service Principal (Read Operations)
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
  // Private Transport \u2014 Delegated User Token (Write Operations)
  // =========================================================================

  /**
   * Makes an authenticated request using a delegated user token.
   * Used for write operations (create/update flows) that require user identity.
   *
   * Token resolution order:
   *   1. Specific userId's token (if provided)
   *   2. Most recently authenticated user's token (auto-resolve)
   *   3. Falls back to service principal (legacy behavior)
   *
   * Includes 401 retry with auto-refresh (same pattern as service principal).
   */
  private async userFlowRequest(
    path: string,
    method: string = 'GET',
    data?: any,
    userId?: string
  ): Promise<any> {
    const url = path.startsWith('http') ? path : `${FLOW_BASE}${path}`;

    // \u2014\u2014\u2014 Attempt 1: Delegated user token \u2014\u2014\u2014
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
          if (resp.status === 204) return { _status: 204, success: true };
          return resp.data;
        } catch (error: any) {
          // 401 \u2192 attempt token refresh + single retry
          if (error.response?.status === 401) {
            console.log(`[Flow] 401 on delegated request for ${resolvedUser}, attempting refresh...`);
            const freshToken = await this.userAuth.getAccessToken(userId);
            if (freshToken && freshToken !== userToken) {
              headers['Authorization'] = `Bearer ${freshToken}`;
              try {
                const retry = await axios({ method, url, data, headers });
                if (retry.status === 204) return { _status: 204, success: true };
                return retry.data;
              } catch (retryError: any) {
                throw this.fmtWriteErr(retryError);
              }
            }
          }
          throw this.fmtWriteErr(error);
        }
      }

      // UserAuth exists but no token available
      console.log('[Flow] No delegated token available for write operation');
    }

    // \u2014\u2014\u2014 Fallback: Service principal \u2014\u2014\u2014
    console.log(`[Flow] ${method} ${url} [service-principal fallback]`);
    console.log('[Flow] \u26a0\ufe0f Write operations may require user auth \u2014 use pa-auth-start if this fails');
    return this.request(url, method, FLOW_SCOPE, data);
  }

  /**
   * Format error for write operations with auth guidance.
   */
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

  /**
   * Resolves the Flow Management API ID from a Dataverse workflowid.
   *
   * ARCHITECTURE BOUNDARY (confirmed by IT team):
   *   Dataverse workflowid \u2260 Flow Management API flowId
   *   Bridge field: workflowidunique (in Dataverse workflow entity)
   *
   * Fallback chain:
   *   1. workflowidunique (primary bridge)
   *   2. resourceid (some environments)
   *   3. Original workflowid (last resort \u2014 may cause 404)
   */
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
  // Flows \u2014 READ Operations (Flow Admin API \u2014 /scopes/admin/ path)
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

  /**
   * Gets detailed information about a specific flow.
   *
   * PROPAGATION AWARENESS (per IT team guidance):
   * After Dataverse creation, the Flow Management API may not see the flow
   * for 5-30 seconds. This method retries on 404 with linear backoff
   * to handle the propagation window transparently.
   */
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

        // Not a 404, or exhausted retries \u2014 throw
        throw error;
      }
    }

    // Should never reach here, but just in case
    throw new Error(`Flow ${flowId} not found after ${PROPAGATION_MAX_RETRIES} retries (propagation timeout)`);
  }

  // =========================================================================
  // Flows \u2014 WRITE Operations (Delegated User Token)
  // =========================================================================
  //
  // ARCHITECTURE NOTES:
  // - Write operations (create/update) use delegated user tokens via Device Code Flow
  // - This ensures flows are owned by the user who created them
  // - If no user is authenticated, falls back to service principal (may fail)
  // - Users authenticate once via pa-auth-start, refresh tokens last ~90 days
  //
  // TOKEN RESOLUTION:
  //   1. userFlowRequest() checks UserAuthManager for a valid delegated token
  //   2. Auto-refreshes if < 5 min remaining (transparent to caller)
  //   3. On 401: refresh + single retry (same pattern as service principal)
  //   4. Falls back to service principal if no user token available
  //
  // =========================================================================

  /**
   * Creates a new Power Automate cloud flow via the Flow Management API.
   * Uses the authenticated user's delegated token to ensure the flow
   * is owned by that user's identity.
   *
   * Requires: User authenticated via pa-auth-start + pa-auth-poll
   */
  async createFlow(
    envId: string,
    displayName: string,
    definition: any,
    state: string = 'Stopped',
    connectionReferences?: any
  ): Promise<any> {
    // Ensure proper Logic Apps schema envelope
    const fullDefinition = definition['$schema']
      ? definition
      : {
          '$schema': WORKFLOW_SCHEMA,
          contentVersion: '1.0.0.0',
          triggers: definition.triggers || {},
          actions: definition.actions || {},
          ...(definition.parameters ? { parameters: definition.parameters } : {}),
        };

    // Build Flow Management API request body (properties envelope)
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
    console.log(`[Flow] State: ${body.properties.state}, Definition: ${JSON.stringify(fullDefinition).length} chars`);

    // POST via delegated user token (or service principal fallback)
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

  /**
   * Updates an existing Power Automate cloud flow via the Flow Management API.
   * Uses the authenticated user's delegated token to ensure proper authorization.
   *
   * Requires: User authenticated via pa-auth-start + pa-auth-poll
   */
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
      body.properties.definition = updates.definition['$schema']
        ? updates.definition
        : {
            '$schema': WORKFLOW_SCHEMA,
            contentVersion: '1.0.0.0',
            triggers: updates.definition.triggers || {},
            actions: updates.definition.actions || {},
            ...(updates.definition.parameters ? { parameters: updates.definition.parameters } : {}),
          };
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

    // PATCH via delegated user token (or service principal fallback)
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
  // Flow Management \u2014 LIFECYCLE (Flow Admin API)
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
