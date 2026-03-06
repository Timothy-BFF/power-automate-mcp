import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';

// =========================================================================
// API Configuration Constants
// =========================================================================
// Service principal (New-PowerAppManagementApp) authentication paths:
// - BAP admin scope for environment listing
// - Flow scope with /scopes/admin/ for flow management (read, enable/disable, delete)
// - Dataverse scope (dynamic per environment) for flow CREATION and UPDATE
// - PowerApps scope for connections listing
//
// CRITICAL: The Flow API (api.flow.microsoft.com) does NOT support creating
// or updating flows with service principal (client_credentials) auth.
// The /scopes/admin/ path only supports GET + lifecycle ops.
// The non-admin path requires delegated (interactive user) auth.
// Solution: Use the Dataverse Web API for flow write operations.
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

// OData headers required for Dataverse Web API requests
const DATAVERSE_HEADERS: Record<string, string> = {
  'Content-Type': 'application/json; charset=utf-8',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  'Prefer': 'return=representation',
};

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

  /**
   * Dataverse Web API request for a specific environment.
   * Discovers the Dataverse instance URL from BAP environment metadata,
   * acquires a dynamically-scoped token ({instanceUrl}/.default),
   * and makes the request with OData headers.
   *
   * This is the ONLY supported path for flow creation/update with service principals.
   */
  private async dataverseRequest(envId: string, path: string, method: string = 'GET', data?: any): Promise<any> {
    const instanceUrl = await this.discoverDataverseUrl(envId);
    if (!instanceUrl) {
      throw new Error(
        `Environment '${envId}' does not have a linked Dataverse instance. ` +
        `Flow creation and update via service principal (application-only auth) requires Dataverse. ` +
        `Please use an environment with Dataverse enabled, or create the flow manually in Power Automate.`
      );
    }

    const scope = `${instanceUrl}/.default`;
    const url = `${instanceUrl}${path}`;
    console.log(`[Dataverse] ${method} ${url}`);

    return this.request(url, method, scope, data, DATAVERSE_HEADERS);
  }

  /**
   * Core HTTP request handler with automatic 401 retry (token refresh).
   * All API transport methods (BAP, Flow, Dataverse, PowerApps) route through here.
   */
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
      // Handle 204 No Content (common for Dataverse PATCH responses)
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
      // Enhanced error logging for debugging
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

  /**
   * Discovers the Dataverse instance URL for a Power Platform environment.
   * Retrieves linkedEnvironmentMetadata.instanceUrl from the BAP Admin API.
   * Results are cached per environment ID to avoid repeated API calls.
   */
  private async discoverDataverseUrl(envId: string): Promise<string | null> {
    if (this.dataverseUrlCache.has(envId)) {
      const cached = this.dataverseUrlCache.get(envId)!;
      console.log(
        `[Dataverse] URL cache ${cached ? 'hit' : 'miss (no Dataverse)'}: ${envId} -> ${cached || 'N/A'}`
      );
      return cached;
    }

    try {
      console.log(`[Dataverse] Discovering instance URL for environment: ${envId}`);
      const envData = await this.bapRequest(
        `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${envId}`
      );

      const instanceUrl = envData?.properties?.linkedEnvironmentMetadata?.instanceUrl;

      if (instanceUrl) {
        // Normalize: remove trailing slash
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
   * Power Automate flows exist in two identity planes simultaneously:
   *   - Dataverse: workflowid (primary key of workflow entity)
   *   - Flow Management API: name/flowId (a separate GUID)
   *
   * The bridge field is 'workflowidunique' in the Dataverse workflow entity.
   * After creating a flow via Dataverse POST, we must query this field
   * to get the ID that the Flow Management API recognizes.
   *
   * Fallback chain:
   *   1. workflowidunique (primary bridge field)
   *   2. resourceid (some environments expose this)
   *   3. Original workflowid (last resort — may cause 404 on Flow API)
   */
  private async resolveFlowApiId(envId: string, dataverseWorkflowId: string): Promise<{ flowApiId: string; dataverseId: string }> {
    try {
      console.log(`[IdResolver] Resolving Flow API ID for Dataverse workflowid: ${dataverseWorkflowId}`);

      const result = await this.dataverseRequest(
        envId,
        `/api/data/v9.2/workflows(${dataverseWorkflowId})?$select=workflowidunique,resourceid,name`,
        'GET'
      );

      const workflowidunique = result?.workflowidunique;
      const resourceid = result?.resourceid;
      const name = result?.name;

      console.log(`[IdResolver] Dataverse response:`);
      console.log(`[IdResolver]   workflowid:       ${dataverseWorkflowId}`);
      console.log(`[IdResolver]   workflowidunique:  ${workflowidunique || '(not found)'}`);
      console.log(`[IdResolver]   resourceid:        ${resourceid || '(not found)'}`);
      console.log(`[IdResolver]   name:              ${name || '(not found)'}`);

      // Use workflowidunique as the Flow API ID (primary bridge)
      // Fall back to resourceid, then to the original workflowid
      const flowApiId = workflowidunique || resourceid || dataverseWorkflowId;

      if (workflowidunique) {
        console.log(`[IdResolver] Resolved via workflowidunique: ${flowApiId}`);
      } else if (resourceid) {
        console.log(`[IdResolver] Resolved via resourceid: ${flowApiId}`);
      } else {
        console.warn(`[IdResolver] No bridge field found — using Dataverse workflowid as fallback: ${flowApiId}`);
      }

      return { flowApiId, dataverseId: dataverseWorkflowId };
    } catch (error: any) {
      console.warn(`[IdResolver] Resolution failed: ${error.message}. Using Dataverse workflowid as fallback.`);
      return { flowApiId: dataverseWorkflowId, dataverseId: dataverseWorkflowId };
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
    return this.flowAdminRequest(
      `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`
    );
  }

  // =========================================================================
  // Flows — WRITE Operations (Dataverse Web API)
  // =========================================================================
  //
  // CRITICAL ARCHITECTURE NOTE:
  // The Power Automate Flow API (api.flow.microsoft.com) does NOT support
  // creating or updating flows with service principal (client_credentials)
  // authentication. This is a confirmed API design limitation:
  //   - /scopes/admin/ paths: Only support GET and lifecycle ops (start/stop/delete)
  //   - Non-admin paths: Require delegated (interactive user) auth
  //
  // The Dataverse Web API (POST/PATCH to /api/data/v9.2/workflows) fully
  // supports application-only auth and is the Microsoft-recommended path
  // for programmatic flow management with service principals.
  //
  // IDENTITY MAPPING (confirmed by IT team investigation):
  //   Dataverse workflowid ≠ Flow Management API flowId
  //   Bridge field: workflowidunique (in Dataverse workflow entity)
  //   After POST creation, a follow-up GET retrieves workflowidunique
  //   which is the ID recognized by the Flow Management API.
  //
  // CLIENTDATA FORMAT (confirmed via Dataverse error diagnostics):
  //   {
  //     "properties": {
  //       "definition": { <Logic Apps schema definition> },
  //       "connectionReferences": {}
  //     },
  //     "schemaVersion": "1.0.0.0"    <-- REQUIRED at root level
  //   }
  //
  //   - displayName must NOT appear inside clientdata — only in entity 'name' field
  //   - schemaVersion is REQUIRED at the clientdata root (not inside properties)
  // =========================================================================

  /**
   * Creates a new Power Automate cloud flow via the Dataverse Web API.
   *
   * Maps to the Dataverse 'workflow' entity with:
   *   - category   = 5  (Modern Flow / Cloud Flow)
   *   - type       = 1  (Definition)
   *   - clientdata = JSON string with {properties: {definition, connectionReferences}, schemaVersion}
   *   - statecode  = 0 (Draft/Stopped) or 1 (Activated/Started)
   *   - statuscode = 1 (Draft) or 2 (Activated)
   *
   * After creation, resolves the Flow Management API ID via workflowidunique
   * so downstream tools (get-flow-details, enable/disable, etc.) work seamlessly.
   */
  async createFlow(
    envId: string,
    displayName: string,
    definition: any,
    state: string = 'Stopped',
    connectionReferences?: any
  ): Promise<any> {
    // Ensure the definition has the proper Logic Apps schema envelope
    const fullDefinition = definition['$schema']
      ? definition
      : {
          '$schema': WORKFLOW_SCHEMA,
          contentVersion: '1.0.0.0',
          triggers: definition.triggers || {},
          actions: definition.actions || {},
          ...(definition.parameters ? { parameters: definition.parameters } : {}),
        };

    // Build the clientdata JSON envelope
    // CRITICAL FORMAT REQUIREMENTS (confirmed via Dataverse error diagnostics):
    //   1. displayName must NOT be inside properties (causes parse failure)
    //   2. schemaVersion MUST be at the root level ("Required property 'schemaVersion' not found")
    //   3. Only definition + connectionReferences go inside properties
    const clientData = {
      properties: {
        definition: fullDefinition,
        connectionReferences: connectionReferences || {},
      },
      schemaVersion: '1.0.0.0',
    };

    // Map Power Automate state to Dataverse statecode/statuscode
    // Stopped -> statecode: 0 (Draft), statuscode: 1
    // Started -> statecode: 1 (Activated), statuscode: 2
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

    const result = await this.dataverseRequest(
      envId,
      '/api/data/v9.2/workflows',
      'POST',
      workflowEntity
    );

    // Extract the Dataverse workflowid from the creation response
    const dataverseWorkflowId = result?.workflowid || 'unknown';
    console.log(`[Dataverse] Flow created. Dataverse workflowid: ${dataverseWorkflowId}`);

    // Resolve the Flow Management API ID via workflowidunique
    // This bridges the Dataverse ↔ Flow API identity boundary
    let flowApiId = dataverseWorkflowId;
    let idSource = 'dataverse-workflowid';

    if (dataverseWorkflowId !== 'unknown') {
      try {
        const resolved = await this.resolveFlowApiId(envId, dataverseWorkflowId);
        flowApiId = resolved.flowApiId;
        idSource = flowApiId !== dataverseWorkflowId ? 'workflowidunique' : 'dataverse-workflowid';
        console.log(`[Dataverse] Flow API ID resolved: ${flowApiId} (via ${idSource})`);
      } catch (resolveErr: any) {
        console.warn(`[Dataverse] ID resolution failed: ${resolveErr.message}. Using Dataverse workflowid.`);
      }
    }

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
        resolvedVia: idSource,
      },
    };
  }

  /**
   * Updates an existing Power Automate cloud flow via the Dataverse Web API.
   * Supports partial updates to displayName, definition, state, and connectionReferences.
   *
   * Note: For state changes only (start/stop), prefer enableDisableFlow() which
   * uses the Flow Admin API and doesn't require Dataverse.
   *
   * IMPORTANT: displayName goes in entity 'name' field ONLY, NOT in clientdata.
   * IMPORTANT: schemaVersion is REQUIRED at the root of clientdata.
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

    // Build updated clientdata if any content properties changed
    // CRITICAL: displayName must NOT be in clientdata — only definition + connectionReferences
    // CRITICAL: schemaVersion MUST be at the root of clientdata
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

      // Only set clientdata if there's actual content to update
      if (Object.keys(clientDataObj.properties).length > 0) {
        patch.clientdata = JSON.stringify(clientDataObj);
        console.log(`[Dataverse] clientdata for update (${patch.clientdata.length} chars)`);
      }
    }

    if (updates.displayName) {
      patch.name = updates.displayName;
    }

    if (updates.state) {
      const isStarted = updates.state === 'Started';
      patch.statecode = isStarted ? 1 : 0;
      patch.statuscode = isStarted ? 2 : 1;
    }

    console.log(`[Dataverse] Updating flow ${flowId} in env ${envId}`);
    console.log(`[Dataverse] Update fields: ${Object.keys(patch).join(', ')}`);

    // Extract bare GUID from flow ID (defensive — handle any format)
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
  // Flow Management — LIFECYCLE (Flow Admin API — /scopes/admin/ path)
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
  // Flow Runs (Flow Admin API — /scopes/admin/ path)
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
