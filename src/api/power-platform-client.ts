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

function ensureRequiredParameters(definition: any, displayName: string): any {
  const def = { ...definition };
  if (!def.parameters) def.parameters = {};
  if (!def.parameters['$connections']) {
    def.parameters['$connections'] = { defaultValue: {}, type: 'Object' };
    console.log(`[Flow] Auto-injected parameters.$connections for '${displayName}'`);
  }
  if (!def.parameters['$authentication']) {
    def.parameters['$authentication'] = { defaultValue: {}, type: 'SecureObject' };
    console.log(`[Flow] Auto-injected parameters.$authentication for '${displayName}'`);
  }
  return def;
}

// =========================================================================
// JSON Repair Engine (v3.0.2)
// =========================================================================

function extractErrorPosition(error: Error): number | null {
  const msg = error.message || '';
  const posMatch = msg.match(/(?:position|column|character)\s+(\d+)/i);
  if (posMatch) return parseInt(posMatch[1], 10);
  return null;
}

function balanceBrackets(str: string): string {
  const chars = Array.from(str);
  const toRemove = new Set<number>();
  const squareStack: number[] = [];
  const curlyStack: number[] = [];
  let inString = false;
  let escapeNext = false;

  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (escapeNext) { escapeNext = false; continue; }
    if (c === '\\' && inString) { escapeNext = true; continue; }
    if (c === '"' && !escapeNext) { inString = !inString; continue; }
    if (inString) continue;
    if (c === '[') squareStack.push(i);
    else if (c === ']') { if (squareStack.length > 0) squareStack.pop(); else toRemove.add(i); }
    else if (c === '{') curlyStack.push(i);
    else if (c === '}') { if (curlyStack.length > 0) curlyStack.pop(); else toRemove.add(i); }
  }

  if (toRemove.size === 0) return str;
  const result: string[] = [];
  for (let i = 0; i < chars.length; i++) {
    if (!toRemove.has(i)) result.push(chars[i]);
  }
  return result.join('');
}

function repairJson(str: string, label: string): any | null {
  console.log(`[Flow:Repair] Attempting JSON repair for '${label}' (${str.length} chars)`);
  let repaired = str;
  let lastPos = -1;
  for (let attempt = 0; attempt < 20; attempt++) {
    try {
      const result = JSON.parse(repaired);
      console.log(`[Flow:Repair] \u2705 Position-targeted repair succeeded after ${attempt} removal(s) for '${label}'`);
      return result;
    } catch (e: any) {
      const pos = extractErrorPosition(e);
      if (pos === null || pos < 0 || pos >= repaired.length) break;
      if (pos === lastPos) break;
      lastPos = pos;
      const badChar = repaired[pos];
      const context = repaired.substring(Math.max(0, pos - 30), Math.min(repaired.length, pos + 30));
      console.log(`[Flow:Repair] Error at pos ${pos}, char '${badChar}' (attempt ${attempt + 1})`);
      console.log(`[Flow:Repair] Context: ...${context}...`);
      repaired = repaired.substring(0, pos) + repaired.substring(pos + 1);
    }
  }

  try {
    const balanced = balanceBrackets(str);
    if (balanced !== str) {
      const result = JSON.parse(balanced);
      console.log(`[Flow:Repair] \u2705 Bracket balancing succeeded for '${label}'`);
      return result;
    }
  } catch (e: any) { /* fall through */ }

  try {
    const combined = balanceBrackets(repaired);
    if (combined !== repaired) {
      const result = JSON.parse(combined);
      console.log(`[Flow:Repair] \u2705 Combined repair succeeded for '${label}'`);
      return result;
    }
  } catch (e: any) { /* fall through */ }

  if (repaired !== str) {
    try { return JSON.parse(repaired); } catch (e: any) { /* fall through */ }
  }

  console.error(`[Flow:Repair] \u274c All repair strategies failed for '${label}'`);
  return null;
}

function aggressiveJsonParse(str: string, label: string): any | null {
  try { return JSON.parse(str); } catch (e) { /* continue */ }

  const cleaned = str.replace(/^\uFEFF/, '').replace(/\0/g, '').trim();
  if (cleaned !== str) {
    try { return JSON.parse(cleaned); } catch (e) { /* continue */ }
  }

  try {
    const inner = JSON.parse(cleaned);
    if (typeof inner === 'string') return JSON.parse(inner);
  } catch (e) { /* continue */ }

  try {
    let val: any = cleaned;
    for (let depth = 0; depth < 5; depth++) {
      val = JSON.parse(val);
      if (typeof val === 'object' && val !== null) return val;
      if (typeof val !== 'string') break;
    }
  } catch (e) { /* continue */ }

  const repaired = repairJson(cleaned, label);
  if (repaired !== null && typeof repaired === 'object') return repaired;

  const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try { return JSON.parse(jsonMatch[0]); } catch (e) { /* continue */ }
    const repairedExtract = repairJson(jsonMatch[0], `${label} (regex-extracted)`);
    if (repairedExtract !== null && typeof repairedExtract === 'object') return repairedExtract;
  }

  const unescaped = cleaned.replace(/\\\\/g, '\\').replace(/\\"/g, '"').replace(/\\n/g, '\n').replace(/\\t/g, '\t');
  if (unescaped !== cleaned) {
    try { return JSON.parse(unescaped); } catch (e) { /* continue */ }
    const repairedUnescaped = repairJson(unescaped, `${label} (unescaped)`);
    if (repairedUnescaped !== null && typeof repairedUnescaped === 'object') return repairedUnescaped;
  }

  console.error(`[Flow:Parse] \u274c ALL parse+repair attempts failed for '${label}'.`);
  return null;
}

function sanitizeDefinition(definition: any, label: string): any {
  let def = definition;

  if (typeof def === 'string') {
    const parsed = aggressiveJsonParse(def, label);
    if (parsed !== null && typeof parsed === 'object') {
      def = parsed;
    } else {
      throw new Error(
        `Cannot parse flow definition: the JSON is malformed. ` +
        `Definition was ${def.length} characters. The tool attempted 7 parse strategies ` +
        `including JSON repair but could not recover a valid object.`
      );
    }
  }

  if (Array.isArray(def)) {
    if (def.length === 1 && def[0] && typeof def[0] === 'object') { def = def[0]; }
    else if (def.length > 1) { def = def[0]; }
    else { throw new Error('Invalid flow definition: received empty array.'); }
  }

  if (def && typeof def === 'object' && !Array.isArray(def)) {
    const numericKeys = Object.keys(def).filter(k => /^\d+$/.test(k));
    if (numericKeys.length > 0) {
      const nonNumericKeys = Object.keys(def).filter(k => !/^\d+$/.test(k));
      if (nonNumericKeys.length === 0 && numericKeys.length === 1) {
        const inner = def[numericKeys[0]];
        if (inner && typeof inner === 'object') def = inner;
      } else if (nonNumericKeys.length === 0 && numericKeys.length > 1) {
        throw new Error(`Invalid flow definition: object has only numeric keys.`);
      } else {
        const cleaned: any = {};
        for (const key of Object.keys(def)) {
          if (!/^\d+$/.test(key)) cleaned[key] = def[key];
        }
        def = cleaned;
      }
    }
  }

  return def;
}

function extractFirstErrorPosition(str: string): number | null {
  try { JSON.parse(str); return null; } catch (e: any) { return extractErrorPosition(e); }
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
  // Private Transport -- Service Principal (Read Operations)
  // =========================================================================

  private async bapRequest(path: string, method: string = 'GET', data?: any): Promise<any> {
    const sep = path.includes('?') ? '&' : '?';
    const url = `${BAP_BASE}${path}${sep}api-version=${BAP_API_VER}`;
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
    return this.request(url, method, POWERAPPS_SCOPE, data);
  }

  private async dataverseRequest(envId: string, path: string, method: string = 'GET', data?: any): Promise<any> {
    const instanceUrl = await this.discoverDataverseUrl(envId);
    if (!instanceUrl) {
      throw new Error(`Environment '${envId}' does not have a linked Dataverse instance.`);
    }
    const scope = `${instanceUrl}/.default`;
    const url = `${instanceUrl}${path}`;
    return this.request(url, method, scope, data, DATAVERSE_HEADERS);
  }

  private async request(
    url: string, method: string, scope: string,
    data?: any, extraHeaders?: Record<string, string>
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
        return guidMatch ? { workflowid: guidMatch[1], _status: 204 } : { _status: 204, success: true };
      }
      return resp.data;
    } catch (error: any) {
      if (error.response?.status === 401) {
        this.tm.invalidate(scope);
        const freshToken = await this.tm.getToken(scope);
        headers['Authorization'] = `Bearer ${freshToken}`;
        const retry = await axios({ method, url, data, headers });
        if (retry.status === 204) {
          const entityId = retry.headers?.['odata-entityid'] || '';
          const guidMatch = entityId.match(/\(([0-9a-f-]+)\)/i);
          return guidMatch ? { workflowid: guidMatch[1], _status: 204 } : { _status: 204, success: true };
        }
        return retry.data;
      }
      if (error.response) {
        console.error(`[API] Error ${error.response.status} on ${method} ${url}`);
        console.error(`[API] Response body:`, JSON.stringify(error.response.data, null, 2));
      }
      throw this.fmtErr(error);
    }
  }

  private fmtErr(error: any): Error {
    if (error.response) {
      const s = error.response.status;
      const d = error.response.data;
      const m = d?.error?.message || d?.error?.code || d?.message || JSON.stringify(d);
      return new Error(`API Error ${s}: ${m}`);
    }
    return new Error(`Request failed: ${error.message}`);
  }

  // =========================================================================
  // Private Transport -- Delegated User Token (Write Operations)
  // =========================================================================

  private async userFlowRequest(
    path: string, method: string = 'GET', data?: any, userId?: string
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
          console.log(`[Flow] \u2705 ${method} ${url} -> ${resp.status} [delegated: ${resolvedUser}]`);
          if (resp.status === 204) return { _status: 204, success: true };
          return resp.data;
        } catch (error: any) {
          if (error.response) {
            console.error(`[Flow] \u274c Error ${error.response.status} on ${method} ${url} [delegated: ${resolvedUser}]`);
            console.error(`[Flow] Response body:`, JSON.stringify(error.response.data, null, 2));
          }
          if (error.response?.status === 401) {
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
    }

    console.log(`[Flow] ${method} ${url} [service-principal fallback]`);
    return this.request(url, method, FLOW_SCOPE, data);
  }

  private fmtWriteErr(error: any): Error {
    if (error.response) {
      const s = error.response.status;
      const d = error.response.data;
      const m = d?.error?.message || d?.error?.code || d?.message || JSON.stringify(d);
      if (s === 401 || s === 403) {
        return new Error(`API Error ${s}: ${m}\n\nUse pa-auth-start to authenticate, then try again.`);
      }
      return new Error(`API Error ${s}: ${m}`);
    }
    return new Error(`Request failed: ${error.message}`);
  }

  // =========================================================================
  // Dataverse URL Discovery
  // =========================================================================

  private async discoverDataverseUrl(envId: string): Promise<string | null> {
    if (this.dataverseUrlCache.has(envId)) return this.dataverseUrlCache.get(envId)!;
    try {
      const envData = await this.bapRequest(
        `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${envId}`
      );
      const instanceUrl = envData?.properties?.linkedEnvironmentMetadata?.instanceUrl;
      if (instanceUrl) {
        const normalized = instanceUrl.replace(/\/+$/, '');
        this.dataverseUrlCache.set(envId, normalized);
        return normalized;
      }
      this.dataverseUrlCache.set(envId, null);
      return null;
    } catch (error: any) {
      this.dataverseUrlCache.set(envId, null);
      return null;
    }
  }

  // =========================================================================
  // Dataverse ID Resolution
  // =========================================================================

  private async resolveFlowApiId(
    envId: string, dataverseWorkflowId: string
  ): Promise<{ flowApiId: string; dataverseId: string; resolvedVia: string }> {
    try {
      const result = await this.dataverseRequest(
        envId,
        `/api/data/v9.2/workflows(${dataverseWorkflowId})?$select=workflowidunique,resourceid,name`,
        'GET'
      );
      if (result?.workflowidunique) {
        return { flowApiId: result.workflowidunique, dataverseId: dataverseWorkflowId, resolvedVia: 'workflowidunique' };
      }
      if (result?.resourceid) {
        return { flowApiId: result.resourceid, dataverseId: dataverseWorkflowId, resolvedVia: 'resourceid' };
      }
      return { flowApiId: dataverseWorkflowId, dataverseId: dataverseWorkflowId, resolvedVia: 'dataverse-workflowid' };
    } catch (error: any) {
      return { flowApiId: dataverseWorkflowId, dataverseId: dataverseWorkflowId, resolvedVia: 'dataverse-workflowid' };
    }
  }

  // =========================================================================
  // Environments
  // =========================================================================

  async listEnvironments(): Promise<any> {
    return this.bapRequest('/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments');
  }

  // =========================================================================
  // Flows -- READ
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

  // =========================================================================
  // GET FLOW DETAILS (v3.0.3 -- dual-path GET)
  //
  // v3.0.2 ROOT CAUSE (confirmed 2026-03-22 via Jose's agent logs):
  //   pa-create-flow writes via delegated user POST to /environments/
  //   pa-get-flow-details reads via admin GET to /scopes/admin/
  //   Admin GET does NOT return definition for delegated-created flows.
  //
  // v3.0.3 FIX: Dual-path GET
  //   1. If delegated token available -> user GET (full definition)
  //   2. Fallback -> admin GET (shell only, with warning)
  // =========================================================================

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    let raw: any = null;
    let fetchPath: string = 'unknown';

    // -----------------------------------------------------------------
    // PATH 1: Delegated user GET (returns full definition)
    // -----------------------------------------------------------------
    if (this.userAuth) {
      const userToken = await this.userAuth.getAccessToken();
      if (userToken) {
        const resolvedUser = this.userAuth.getDefaultUserId() || 'delegated-user';
        const userUrl = `${FLOW_BASE}/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`;
        console.log(`[Flow:Details] Trying delegated GET for flow ${flowId} [user: ${resolvedUser}]`);
        console.log(`[Flow:Details] GET ${userUrl}`);

        try {
          const headers: Record<string, string> = {
            'Authorization': `Bearer ${userToken}`,
            'Content-Type': 'application/json',
          };
          const resp = await axios({ method: 'GET', url: userUrl, headers });
          raw = resp.data;
          fetchPath = `delegated (${resolvedUser})`;
          console.log(`[Flow:Details] \u2705 Delegated GET succeeded for flow ${flowId} [user: ${resolvedUser}]`);
        } catch (delegatedErr: any) {
          const status = delegatedErr.response?.status;
          console.log(`[Flow:Details] Delegated GET failed: ${status || delegatedErr.message}. Falling back to admin GET.`);

          // If 401, try refreshing once
          if (status === 401) {
            console.log(`[Flow:Details] 401 on delegated GET, attempting token refresh...`);
            const freshToken = await this.userAuth.getAccessToken();
            if (freshToken && freshToken !== userToken) {
              try {
                const retryHeaders: Record<string, string> = {
                  'Authorization': `Bearer ${freshToken}`,
                  'Content-Type': 'application/json',
                };
                const retryResp = await axios({ method: 'GET', url: userUrl, headers: retryHeaders });
                raw = retryResp.data;
                fetchPath = `delegated-retry (${resolvedUser})`;
                console.log(`[Flow:Details] \u2705 Delegated GET retry succeeded for flow ${flowId}`);
              } catch (retryErr: any) {
                console.log(`[Flow:Details] Delegated GET retry also failed: ${retryErr.response?.status || retryErr.message}`);
              }
            }
          }
        }
      } else {
        console.log(`[Flow:Details] No delegated user token available, using admin GET`);
      }
    }

    // -----------------------------------------------------------------
    // PATH 2: Admin GET fallback (may omit definition body)
    // -----------------------------------------------------------------
    if (!raw) {
      const adminUrl = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`;
      console.log(`[Flow:Details] Using admin GET for flow ${flowId}`);

      for (let attempt = 0; attempt <= PROPAGATION_MAX_RETRIES; attempt++) {
        try {
          raw = await this.flowAdminRequest(adminUrl);
          fetchPath = 'admin (service-principal)';
          if (attempt > 0) {
            console.log(`[Flow] Flow ${flowId} found after ${attempt} retries`);
          }
          break;
        } catch (error: any) {
          const is404 = error.message?.includes('404') || error.message?.includes('Could not find flow');
          if (is404 && attempt < PROPAGATION_MAX_RETRIES) {
            const delayMs = PROPAGATION_RETRY_DELAY_MS * (attempt + 1);
            console.log(`[Flow] 404 on attempt ${attempt + 1}, waiting ${delayMs}ms...`);
            await sleep(delayMs);
            continue;
          }
          throw error;
        }
      }
    }

    if (!raw) {
      throw new Error(`Flow ${flowId} not found after ${PROPAGATION_MAX_RETRIES} retries`);
    }

    // =================================================================
    // DEFINITION EXTRACTION (v3.0.3)
    // =================================================================
    const props = raw?.properties || {};
    const definition = props?.definition || {};
    const triggers = definition?.triggers || {};
    const actions = definition?.actions || {};
    const connectionRefs = props?.connectionReferences || {};

    const triggerNames = Object.keys(triggers);
    const actionNames = Object.keys(actions);
    const connectionNames = Object.keys(connectionRefs);

    console.log(`[Flow:Details] Flow ${flowId} via ${fetchPath}: ${triggerNames.length} trigger(s), ${actionNames.length} action(s), ${connectionNames.length} connection(s)`);
    if (triggerNames.length > 0) console.log(`[Flow:Details] Triggers: [${triggerNames.join(', ')}]`);
    if (actionNames.length > 0) console.log(`[Flow:Details] Actions: [${actionNames.join(', ')}]`);
    if (triggerNames.length === 0 && actionNames.length === 0) {
      console.warn(`[Flow:Details] \u26a0\ufe0f Definition appears empty for flow ${flowId} (fetched via ${fetchPath}).`);
    }

    return {
      flowId: raw?.name || flowId,
      displayName: props?.displayName || 'Unknown',
      state: props?.state || 'Unknown',
      createdTime: props?.createdTime,
      lastModifiedTime: props?.lastModifiedTime,
      creator: props?.creator?.userId || props?.creator?.objectId || 'unknown',

      definition: {
        schema: definition?.['$schema'] || 'none',
        contentVersion: definition?.contentVersion || 'none',
        triggerCount: triggerNames.length,
        triggerNames: triggerNames,
        actionCount: actionNames.length,
        actionNames: actionNames,
        hasParameters: !!(definition?.parameters),
        parameterNames: definition?.parameters ? Object.keys(definition.parameters) : [],
        triggers: triggers,
        actions: actions,
      },

      connectionReferences: connectionRefs,
      connectionCount: connectionNames.length,
      connectionNames: connectionNames,

      _fetchedVia: fetchPath,

      _definitionStatus: (triggerNames.length > 0 || actionNames.length > 0)
        ? 'POPULATED'
        : 'EMPTY_OR_NOT_RETURNED',
      _definitionNote: (triggerNames.length === 0 && actionNames.length === 0)
        ? `WARNING: Definition appears empty (fetched via ${fetchPath}). ` +
          'The admin API sometimes omits the definition body. ' +
          'If a previous create or PATCH returned 200, the definition IS saved. ' +
          'Check the flow in Power Automate portal to confirm. ' +
          'Do NOT delete and recreate based solely on an empty API response.'
        : `Flow has ${triggerNames.length} trigger(s) [${triggerNames.join(', ')}] and ${actionNames.length} action(s) [${actionNames.join(', ')}]. Fetched via ${fetchPath}.`,

      _raw: raw,
    };
  }

  // =========================================================================
  // Flows -- WRITE Operations (Delegated User Token)
  // =========================================================================

  async createFlow(
    envId: string, displayName: string, definition: any,
    state: string = 'Stopped', connectionReferences?: any
  ): Promise<any> {
    definition = sanitizeDefinition(definition, displayName);
    let fullDefinition = { ...definition };
    if (!fullDefinition['$schema']) fullDefinition['$schema'] = WORKFLOW_SCHEMA;
    if (!fullDefinition.contentVersion) fullDefinition.contentVersion = '1.0.0.0';
    if (!fullDefinition.triggers) fullDefinition.triggers = {};
    if (!fullDefinition.actions) fullDefinition.actions = {};

    fullDefinition = ensureRequiredParameters(fullDefinition, displayName);

    const triggerCount = Object.keys(fullDefinition.triggers).length;
    const actionCount = Object.keys(fullDefinition.actions).length;

    if (triggerCount === 0 && actionCount === 0) {
      throw new Error('Invalid flow definition: both triggers and actions are empty.');
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

    console.log(`[Flow] Creating flow '${displayName}' in env ${envId}`);
    console.log(`[Flow] State: ${body.properties.state}, Definition: ${JSON.stringify(fullDefinition).length} chars, Triggers: ${triggerCount}, Actions: ${actionCount}`);

    const result = await this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows?api-version=${FLOW_API_VER}`,
      'POST', body
    );

    const newFlowId = result?.name || 'unknown';
    console.log(`[Flow] \u2705 Flow created: ${newFlowId} (${displayName})`);

    return {
      status: 'created',
      flowId: newFlowId,
      name: newFlowId,
      displayName: result?.properties?.displayName || displayName,
      state: result?.properties?.state || state,
      definition: fullDefinition,
      connectionReferences: connectionReferences || {},
      _source: 'flow-management-api',
      _authType: this.userAuth?.hasAuthenticatedUser() ? 'delegated' : 'service-principal',
      _note: 'Created via Flow Management API. If the create returned 200/201, the definition IS saved.',
    };
  }

  async updateFlow(
    envId: string, flowId: string,
    updates: { displayName?: string; definition?: any; state?: string; connectionReferences?: any; }
  ): Promise<any> {
    const body: any = { properties: {} };

    if (updates.displayName) body.properties.displayName = updates.displayName;

    if (updates.definition) {
      let def = sanitizeDefinition(updates.definition, `flow ${flowId}`);
      def = { ...def };
      if (!def['$schema']) def['$schema'] = WORKFLOW_SCHEMA;
      if (!def.contentVersion) def.contentVersion = '1.0.0.0';
      if (!def.triggers) def.triggers = {};
      if (!def.actions) def.actions = {};
      def = ensureRequiredParameters(def, `flow ${flowId}`);
      body.properties.definition = def;
    }

    if (updates.connectionReferences) body.properties.connectionReferences = updates.connectionReferences;
    if (updates.state) body.properties.state = updates.state === 'Started' ? 'Started' : 'Stopped';

    console.log(`[Flow] Updating flow ${flowId}: ${Object.keys(body.properties).join(', ')}`);

    const result = await this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`,
      'PATCH', body
    );

    console.log(`[Flow] \u2705 Flow updated: ${flowId}`);

    return {
      name: result?.name || flowId,
      properties: {
        ...(updates.displayName ? { displayName: result?.properties?.displayName || updates.displayName } : {}),
        ...(updates.state ? { state: result?.properties?.state || updates.state } : {}),
        ...(updates.definition ? { definition: updates.definition } : {}),
      },
      _source: 'flow-management-api',
      _authType: this.userAuth?.hasAuthenticatedUser() ? 'delegated' : 'service-principal',
      _raw: result,
    };
  }

  // =========================================================================
  // Flow Lifecycle
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

  // =========================================================================
  // TRIGGER FLOW (v3.0.2 -- delegated user path)
  // =========================================================================

  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    console.log(`[Flow] Triggering flow ${flowId} via delegated user path`);
    return this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/triggers/manual/run?api-version=${FLOW_API_VER}`,
      'POST', body || {}
    );
  }

  // =========================================================================
  // Flow Runs
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
  // Connections
  // =========================================================================

  async listConnections(envId: string): Promise<any> {
    return this.powerAppsAdminRequest(
      `/providers/Microsoft.PowerApps/scopes/admin/environments/${envId}/connections`
    );
  }
}
