import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';

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

// Propagation delay settings (Dataverse write → Flow Management API visibility)
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
  private dataverseUrlCache: Map<string, string | null> = new Map();

  constructor(tokenManager: AzureTokenManager) {
    this.tm = tokenManager;
  }

  // =========================================================================
  // Private Transport Methods
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
   *   Dataverse workflowid ≠ Flow Management API flowId
   *   Bridge field: workflowidunique (in Dataverse workflow entity)
   *
   * Fallback chain:
   *   1. workflowidunique (primary bridge)
   *   2. resourceid (some environments)
   *   3. Original workflowid (last resort — may cause 404)
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

      console.warn(`[IdResolver] No bridge field found — using Dataverse workflowid as fallback`);
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

        // Not a 404, or exhausted retries — throw
        throw error;
      }
    }

    // Should never reach here, but just in case
    throw new Error(`Flow ${flowId} not found after ${PROPAGATION_MAX_RETRIES} retries (propagation timeout)`);
  }

  // =========================================================================
  // Flows — WRITE Operations (Dataverse Web API)
  // =========================================================================
  //
  // ARCHITECTURE NOTES:
  // - Flow API does NOT support create/update with service principal auth
  // - Dataverse Web API is the supported path for application-only auth
  // - Identity mapping: workflowidunique bridges Dataverse ↔ Flow API
  // - Propagation delay: 5-30s between Dataverse write and Flow API visibility
  //
  // CLIENTDATA FORMAT:
  //   {
  //     "properties": { "definition": {...}, "connectionReferences": {} },
  //     "schemaVersion": "1.0.0.0"
  //   }
  //   - displayName in entity 'name' field ONLY
  //   - schemaVersion REQUIRED at clientdata root
  // =========================================================================

  /**
   * Creates a new Power Automate cloud flow via the Dataverse Web API.
   *
   * Post-creation workflow (per IT team recommendation):
   *   1. POST workflow entity → get Dataverse workflowid
   *   2. GET workflowidunique → resolve Flow API ID
   *   3. Wait for propagation delay (5s)
   *   4. Verify flow visible via Flow Management API
   *   5. Return with full ID mapping + propagation status
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

    // Build clientdata (schemaVersion at root, displayName NOT inside)
    const clientData = {
      properties: {
        definition: fullDefinition,
        connectionReferences: connectionReferences || {},
      },
      schemaVersion: '1.0.0.0',
    };

    const isStarted = state === 'Started';
    const clientDataStr = JSON.stringify(clientData);

    const workflowEntity = {
      name: displayName,
      type: 1,
      category: 5,
      statecode: isStarted ? 1 : 0,
      statuscode: isStarted ? 2 : 1,
      primaryentity: 'none',
      clientdata: clientDataStr,
    };

    console.log(`[Dataverse] Creating flow: "${displayName}" in env ${envId} (state: ${state})`);
    console.log(`[Dataverse] Workflow entity: category=5, type=1, statecode=${workflowEntity.statecode}`);
    console.log(`[Dataverse] Definition triggers: ${Object.keys(fullDefinition.triggers || {}).join(', ') || '(none)'}`);
    console.log(`[Dataverse] Definition actions: ${Object.keys(fullDefinition.actions || {}).join(', ') || '(none)'}`);
    console.log(`[Dataverse] clientdata length: ${clientDataStr.length} chars`);

    // ---- Step 1: Create the workflow in Dataverse ----
    const result = await this.dataverseRequest(
      envId,
      '/api/data/v9.2/workflows',
      'POST',
      workflowEntity
    );

    const dataverseWorkflowId = result?.workflowid || 'unknown';
    console.log(`[Dataverse] Flow created. Dataverse workflowid: ${dataverseWorkflowId}`);

    // ---- Step 2: Resolve Flow Management API ID ----
    let flowApiId = dataverseWorkflowId;
    let resolvedVia = 'dataverse-workflowid';

    if (dataverseWorkflowId !== 'unknown') {
      try {
        const resolved = await this.resolveFlowApiId(envId, dataverseWorkflowId);
        flowApiId = resolved.flowApiId;
        resolvedVia = resolved.resolvedVia;
        console.log(`[Dataverse] Flow API ID resolved: ${flowApiId} (via ${resolvedVia})`);
      } catch (resolveErr: any) {
        console.warn(`[Dataverse] ID resolution failed: ${resolveErr.message}. Using Dataverse workflowid.`);
      }
    }

    // ---- Step 3: Propagation delay + verification ----
    let propagationVerified = false;
    let propagationAttempts = 0;
    const propagationStart = Date.now();

    console.log(`[Propagation] Waiting ${PROPAGATION_DELAY_MS}ms for Flow Management API visibility...`);
    await sleep(PROPAGATION_DELAY_MS);

    // One verification attempt via Flow Management API
    try {
      propagationAttempts = 1;
      const verifyUrl = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowApiId}?api-version=${FLOW_API_VER}`;
      await this.flowAdminRequest(verifyUrl);
      propagationVerified = true;
      console.log(`[Propagation] Flow ${flowApiId} verified in Flow Management API after ${Date.now() - propagationStart}ms`);
    } catch (verifyErr: any) {
      const is404 = verifyErr.message?.includes('404') || verifyErr.message?.includes('Could not find flow');
      if (is404) {
        console.log(`[Propagation] Flow ${flowApiId} not yet visible (404). Propagation still in progress. ` +
          `Downstream get-flow-details will retry with backoff.`);
      } else {
        console.warn(`[Propagation] Verification failed (non-404): ${verifyErr.message}`);
      }
    }

    const propagationMs = Date.now() - propagationStart;

    return {
      name: flowApiId,
      id: `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowApiId}`,
      type: 'Microsoft.ProcessSimple/environments/flows',
      properties: {
        displayName,
        state: isStarted ? 'Started' : 'Stopped',
        definition: fullDefinition,
        connectionReferences: connectionReferences || {},
      },
      _source: 'dataverse',
      _idMapping: {
        flowApiId,
        dataverseWorkflowId,
        resolvedVia,
      },
      _propagation: {
        verified: propagationVerified,
        delayMs: propagationMs,
        attempts: propagationAttempts,
        ...(propagationVerified ? { verifiedAt: new Date().toISOString() } : {}),
      },
    };
  }

  /**
   * Updates an existing Power Automate cloud flow via the Dataverse Web API.
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
    const patch: any = {};

    if (updates.definition || updates.connectionReferences) {
      const fullDefinition = updates.definition
        ? (updates.definition['$schema']
            ? updates.definition
            : {
                '$schema': WORKFLOW_SCHEMA,
                contentVersion: '1.0.0.0',
                triggers: updates.definition.triggers || {},
                actions: updates.definition.actions || {},
                ...(updates.definition.parameters ? { parameters: updates.definition.parameters } : {}),
              })
        : undefined;

      const clientDataObj: any = { properties: {}, schemaVersion: '1.0.0.0' };
      if (fullDefinition) clientDataObj.properties.definition = fullDefinition;
      if (updates.connectionReferences) {
        clientDataObj.properties.connectionReferences = updates.connectionReferences;
      }

      if (Object.keys(clientDataObj.properties).length > 0) {
        patch.clientdata = JSON.stringify(clientDataObj);
        console.log(`[Dataverse] clientdata for update (${patch.clientdata.length} chars)`);
      }
    }

    if (updates.displayName) patch.name = updates.displayName;

    if (updates.state) {
      const isStarted = updates.state === 'Started';
      patch.statecode = isStarted ? 1 : 0;
      patch.statuscode = isStarted ? 2 : 1;
    }

    console.log(`[Dataverse] Updating flow ${flowId} in env ${envId}`);
    console.log(`[Dataverse] Update fields: ${Object.keys(patch).join(', ')}`);

    const workflowGuid = flowId.includes('/') ? (flowId.split('/').pop() || flowId) : flowId;

    const result = await this.dataverseRequest(
      envId,
      `/api/data/v9.2/workflows(${workflowGuid})`,
      'PATCH',
      patch
    );

    return {
      name: flowId,
      properties: {
        ...(updates.displayName ? { displayName: updates.displayName } : {}),
        ...(updates.state ? { state: updates.state } : {}),
        ...(updates.definition ? { definition: updates.definition } : {}),
      },
      _source: 'dataverse',
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
