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

// =========================================================================
// JSON Repair Engine (v3.0.2)
//
// ROOT CAUSE (confirmed 2026-03-20 23:50 via diagnostic logs):
// The agent generates syntactically invalid JSON with stray brackets.
// Error: "Unexpected token ] in JSON at position 619"
//
// Strategies:
//   1. Position-targeted iterative removal
//   2. Bracket-aware balancing
//   3. Combined: position removal + bracket balancing
// =========================================================================

function extractErrorPosition(error: Error): number | null {
  const msg = error.message || '';
  const posMatch = msg.match(/(?:position|column|character)\s+(\d+)/i);
  if (posMatch) {
    return parseInt(posMatch[1], 10);
  }
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

    if (escapeNext) {
      escapeNext = false;
      continue;
    }

    if (c === '\\' && inString) {
      escapeNext = true;
      continue;
    }

    if (c === '"' && !escapeNext) {
      inString = !inString;
      continue;
    }

    if (inString) continue;

    if (c === '[') {
      squareStack.push(i);
    } else if (c === ']') {
      if (squareStack.length > 0) {
        squareStack.pop();
      } else {
        toRemove.add(i);
      }
    } else if (c === '{') {
      curlyStack.push(i);
    } else if (c === '}') {
      if (curlyStack.length > 0) {
        curlyStack.pop();
      } else {
        toRemove.add(i);
      }
    }
  }

  if (toRemove.size === 0) {
    return str;
  }

  const result: string[] = [];
  for (let i = 0; i < chars.length; i++) {
    if (!toRemove.has(i)) {
      result.push(chars[i]);
    }
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
      if (pos === null || pos < 0 || pos >= repaired.length) {
        console.log(`[Flow:Repair] Cannot extract position from error: ${e.message}`);
        break;
      }
      if (pos === lastPos) {
        console.log(`[Flow:Repair] Stuck at position ${pos}, switching to bracket balancing`);
        break;
      }
      lastPos = pos;

      const badChar = repaired[pos];
      const context = repaired.substring(Math.max(0, pos - 30), Math.min(repaired.length, pos + 30));
      const contextMarker = ' '.repeat(Math.min(pos, 30)) + '^';
      console.log(`[Flow:Repair] Error at pos ${pos}, char '${badChar}' (attempt ${attempt + 1})`);
      console.log(`[Flow:Repair] Context: ...${context}...`);
      console.log(`[Flow:Repair]          ${contextMarker}`);

      repaired = repaired.substring(0, pos) + repaired.substring(pos + 1);
    }
  }

  console.log(`[Flow:Repair] Trying bracket balancing on original string for '${label}'`);
  try {
    const balanced = balanceBrackets(str);
    if (balanced !== str) {
      const removedCount = str.length - balanced.length;
      console.log(`[Flow:Repair] Bracket balancer removed ${removedCount} unmatched bracket(s)`);
      const result = JSON.parse(balanced);
      console.log(`[Flow:Repair] \u2705 Bracket balancing succeeded for '${label}'`);
      return result;
    } else {
      console.log(`[Flow:Repair] Bracket balancer found no unmatched brackets`);
    }
  } catch (e: any) {
    console.log(`[Flow:Repair] Bracket balancing parse failed: ${e.message}`);
  }

  console.log(`[Flow:Repair] Trying combined repair for '${label}'`);
  try {
    const combined = balanceBrackets(repaired);
    if (combined !== repaired) {
      const removedCount = repaired.length - combined.length;
      console.log(`[Flow:Repair] Combined: removed ${removedCount} additional bracket(s)`);
      const result = JSON.parse(combined);
      console.log(`[Flow:Repair] \u2705 Combined repair succeeded for '${label}'`);
      return result;
    }
  } catch (e: any) {
    console.log(`[Flow:Repair] Combined repair parse failed: ${e.message}`);
  }

  if (repaired !== str) {
    try {
      const result = JSON.parse(repaired);
      console.log(`[Flow:Repair] \u2705 Position-repaired string parsed on second pass for '${label}'`);
      return result;
    } catch (e: any) {
      console.log(`[Flow:Repair] Position-repaired string still invalid: ${e.message}`);
    }
  }

  console.error(`[Flow:Repair] \u274c All repair strategies failed for '${label}'`);
  return null;
}

function aggressiveJsonParse(str: string, label: string): any | null {
  const preview = str.length > 500 ? str.substring(0, 500) + '...' : str;
  console.log(`[Flow:Parse] Raw string (${str.length} chars) for '${label}': ${preview}`);

  try {
    const result = JSON.parse(str);
    console.log(`[Flow:Parse] \u2705 Direct JSON.parse succeeded for '${label}'`);
    return result;
  } catch (e) {
    console.log(`[Flow:Parse] Direct parse failed for '${label}': ${(e as Error).message}`);
  }

  const cleaned = str.replace(/^\uFEFF/, '').replace(/\0/g, '').trim();
  if (cleaned !== str) {
    try {
      const result = JSON.parse(cleaned);
      console.log(`[Flow:Parse] \u2705 Parsed after BOM/whitespace cleanup for '${label}'`);
      return result;
    } catch (e) {
      console.log(`[Flow:Parse] BOM-cleaned parse failed for '${label}'`);
    }
  }

  try {
    const inner = JSON.parse(cleaned);
    if (typeof inner === 'string') {
      const result = JSON.parse(inner);
      console.log(`[Flow:Parse] \u2705 Double-decode succeeded for '${label}'`);
      return result;
    }
  } catch (e) {
    console.log(`[Flow:Parse] Double-decode failed for '${label}'`);
  }

  try {
    let val: any = cleaned;
    for (let depth = 0; depth < 5; depth++) {
      val = JSON.parse(val);
      if (typeof val === 'object' && val !== null) {
        console.log(`[Flow:Parse] \u2705 Multi-decode succeeded at depth ${depth + 1} for '${label}'`);
        return val;
      }
      if (typeof val !== 'string') break;
    }
  } catch (e) {
    console.log(`[Flow:Parse] Multi-decode failed for '${label}'`);
  }

  const repaired = repairJson(cleaned, label);
  if (repaired !== null && typeof repaired === 'object') {
    return repaired;
  }

  const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const result = JSON.parse(jsonMatch[0]);
      console.log(`[Flow:Parse] \u2705 Regex extraction succeeded for '${label}'`);
      return result;
    } catch (e) {
      console.log(`[Flow:Parse] Regex extraction failed for '${label}'`);

      const repairedExtract = repairJson(jsonMatch[0], `${label} (regex-extracted)`);
      if (repairedExtract !== null && typeof repairedExtract === 'object') {
        return repairedExtract;
      }
    }
  }

  const unescaped = cleaned
    .replace(/\\\\/g, '\\')
    .replace(/\\"/g, '"')
    .replace(/\\n/g, '\n')
    .replace(/\\t/g, '\t');
  if (unescaped !== cleaned) {
    try {
      const result = JSON.parse(unescaped);
      console.log(`[Flow:Parse] \u2705 Unescape parse succeeded for '${label}'`);
      return result;
    } catch (e) {
      console.log(`[Flow:Parse] Unescape parse failed for '${label}'`);

      const repairedUnescaped = repairJson(unescaped, `${label} (unescaped)`);
      if (repairedUnescaped !== null && typeof repairedUnescaped === 'object') {
        return repairedUnescaped;
      }
    }
  }

  console.error(`[Flow:Parse] \u274c ALL parse+repair attempts failed for '${label}'.`);
  console.error(`[Flow:Parse] String char codes (first 50): [${Array.from(str.substring(0, 50)).map(c => c.charCodeAt(0)).join(', ')}]`);
  return null;
}

function sanitizeDefinition(definition: any, label: string): any {
  const inputType = typeof definition;
  const inputIsArray = Array.isArray(definition);
  const inputKeys = (definition && typeof definition === 'object' && !inputIsArray)
    ? Object.keys(definition)
    : null;
  console.log(`[Flow:Sanitize] Input for '${label}': type=${inputType}, isArray=${inputIsArray}, keys=${inputKeys ? JSON.stringify(inputKeys.slice(0, 20)) : 'N/A'}${inputKeys && inputKeys.length > 20 ? ` (+${inputKeys.length - 20} more)` : ''}`);

  let def = definition;

  if (typeof def === 'string') {
    const parsed = aggressiveJsonParse(def, label);
    if (parsed !== null && typeof parsed === 'object') {
      def = parsed;
    } else {
      const errorPos = extractFirstErrorPosition(def);
      const contextHint = errorPos !== null
        ? ` Parse error near position ${errorPos}: "...${def.substring(Math.max(0, errorPos - 20), Math.min(def.length, errorPos + 20))}..."`
        : '';
      throw new Error(
        `Cannot parse flow definition: the JSON is malformed.${contextHint} ` +
        `This typically happens when Power Automate expressions (e.g., @{...}, @addDays(...)) ` +
        `contain unescaped brackets. Please ensure all expressions are properly enclosed in ` +
        `double-quoted JSON strings and that brackets within string values are balanced. ` +
        `Definition was ${def.length} characters. The tool attempted 7 parse strategies ` +
        `including JSON repair but could not recover a valid object.`
      );
    }
  }

  if (Array.isArray(def)) {
    if (def.length === 1 && def[0] && typeof def[0] === 'object') {
      console.log(`[Flow:Sanitize] Unwrapped real array for '${label}'`);
      def = def[0];
    } else if (def.length > 1) {
      console.warn(`[Flow:Sanitize] Definition is array with ${def.length} elements for '${label}' - using first element`);
      def = def[0];
    } else {
      throw new Error('Invalid flow definition: received empty array.');
    }
  }

  if (def && typeof def === 'object' && !Array.isArray(def)) {
    const numericKeys = Object.keys(def).filter(k => /^\d+$/.test(k));

    if (numericKeys.length > 0) {
      console.log(`[Flow:Sanitize] Stripping ${numericKeys.length} numeric key(s) from '${label}': [${numericKeys.slice(0, 10).join(', ')}]${numericKeys.length > 10 ? '...' : ''}`);

      const nonNumericKeys = Object.keys(def).filter(k => !/^\d+$/.test(k));
      if (nonNumericKeys.length === 0 && numericKeys.length === 1) {
        const inner = def[numericKeys[0]];
        if (inner && typeof inner === 'object') {
          console.log(`[Flow:Sanitize] Pure numeric-key object - unwrapping '${numericKeys[0]}' for '${label}'`);
          def = inner;
        }
      } else if (nonNumericKeys.length === 0 && numericKeys.length > 1) {
        throw new Error(
          `Invalid flow definition: object has only numeric keys (${numericKeys.length} chars) ` +
          `which indicates the definition string was spread into characters. ` +
          `Please provide a valid JSON object for the definition.`
        );
      } else {
        const cleaned: any = {};
        for (const key of Object.keys(def)) {
          if (!/^\d+$/.test(key)) {
            cleaned[key] = def[key];
          }
        }
        def = cleaned;
      }
    }
  }

  const outputKeys = (def && typeof def === 'object' && !Array.isArray(def))
    ? Object.keys(def)
    : null;
  console.log(`[Flow:Sanitize] Output for '${label}': keys=${outputKeys ? JSON.stringify(outputKeys.slice(0, 20)) : 'N/A'}`);

  return def;
}

function extractFirstErrorPosition(str: string): number | null {
  try {
    JSON.parse(str);
    return null;
  } catch (e: any) {
    return extractErrorPosition(e);
  }
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
  // Private Transport Methods -- Service Principal (Read Operations)
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
  // Private Transport -- Delegated User Token (Write Operations)
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
    console.log('[Flow] \u26a0\ufe0f Write operations may require user auth - use pa-auth-start if this fails');
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

      console.warn(`[IdResolver] No bridge field found - using Dataverse workflowid as fallback`);
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
  // Flows -- READ Operations (Flow Admin API -- /scopes/admin/ path)
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
  // GET FLOW DETAILS (v3.0.2 -- structured response)
  //
  // ROOT CAUSE (confirmed 2026-03-21 via Jose's agent logs):
  //   The Flow Admin API returns the definition nested inside
  //   response.properties.definition. Agents reading the top-level
  //   response object see no triggers/actions and incorrectly conclude
  //   "definition is empty" -- leading to unnecessary delete+recreate
  //   cycles and the false conclusion "MCP tool cannot push definitions".
  //
  // FIX: Surface definition contents (trigger names, action names,
  //   counts, connection references) at the top level of the return
  //   object. Include _definitionStatus indicator and _definitionNote
  //   guidance to prevent agents from making destructive decisions
  //   based on a misread API response.
  // =========================================================================

  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    const url = `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows/${flowId}?api-version=${FLOW_API_VER}`;

    let raw: any = null;
    for (let attempt = 0; attempt <= PROPAGATION_MAX_RETRIES; attempt++) {
      try {
        raw = await this.flowAdminRequest(url);
        if (attempt > 0) {
          console.log(`[Flow] Flow ${flowId} found after ${attempt} retries (propagation delay)`);
        }
        break;
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

    if (!raw) {
      throw new Error(`Flow ${flowId} not found after ${PROPAGATION_MAX_RETRIES} retries (propagation timeout)`);
    }

    // =================================================================
    // DEFINITION EXTRACTION (v3.0.2)
    //
    // The admin GET API nests data inside properties.definition.
    // We surface everything at the top level so agents cannot misread.
    // =================================================================
    const props = raw?.properties || {};
    const definition = props?.definition || {};
    const triggers = definition?.triggers || {};
    const actions = definition?.actions || {};
    const connectionRefs = props?.connectionReferences || {};

    const triggerNames = Object.keys(triggers);
    const actionNames = Object.keys(actions);
    const connectionNames = Object.keys(connectionRefs);

    // Diagnostic logging
    console.log(`[Flow:Details] Flow ${flowId}: ${triggerNames.length} trigger(s), ${actionNames.length} action(s), ${connectionNames.length} connection(s)`);
    if (triggerNames.length > 0) {
      console.log(`[Flow:Details] Triggers: [${triggerNames.join(', ')}]`);
    }
    if (actionNames.length > 0) {
      console.log(`[Flow:Details] Actions: [${actionNames.join(', ')}]`);
    }
    if (triggerNames.length === 0 && actionNames.length === 0) {
      console.warn(`[Flow:Details] \u26a0\ufe0f Definition appears empty for flow ${flowId}. ` +
        `This may be normal if the admin GET endpoint omits the definition body. ` +
        `Check the flow in the Power Automate portal to confirm.`);
    }

    // =================================================================
    // STRUCTURED RETURN
    //
    // Agent sees everything at the top level -- no digging required.
    // _definitionStatus and _definitionNote provide guardrails to
    // prevent agents from making destructive decisions (delete/recreate)
    // based on a misread API response.
    // =================================================================
    return {
      flowId: raw?.name || flowId,
      displayName: props?.displayName || 'Unknown',
      state: props?.state || 'Unknown',
      createdTime: props?.createdTime,
      lastModifiedTime: props?.lastModifiedTime,
      creator: props?.creator?.userId || props?.creator?.objectId || 'unknown',

      // === DEFINITION (surfaced clearly for agent consumption) ===
      definition: {
        schema: definition?.['$schema'] || 'none',
        contentVersion: definition?.contentVersion || 'none',
        triggerCount: triggerNames.length,
        triggerNames: triggerNames,
        actionCount: actionNames.length,
        actionNames: actionNames,
        hasParameters: !!(definition?.parameters),
        parameterNames: definition?.parameters ? Object.keys(definition.parameters) : [],
        // Full trigger and action objects for detailed inspection
        triggers: triggers,
        actions: actions,
      },

      // === CONNECTIONS ===
      connectionReferences: connectionRefs,
      connectionCount: connectionNames.length,
      connectionNames: connectionNames,

      // === DEFINITION STATUS -- critical for agent decision-making ===
      _definitionStatus: (triggerNames.length > 0 || actionNames.length > 0)
        ? 'POPULATED'
        : 'EMPTY_OR_NOT_RETURNED',
      _definitionNote: (triggerNames.length === 0 && actionNames.length === 0)
        ? 'WARNING: The admin API sometimes omits the full definition body in GET responses. ' +
          'An empty definition here does NOT mean the flow has no definition. ' +
          'If a previous create or PATCH returned 200, the definition IS saved. ' +
          'Check the flow in the Power Automate portal designer to confirm visually. ' +
          'Do NOT delete and recreate the flow based solely on an empty API response.'
        : `Flow has ${triggerNames.length} trigger(s) [${triggerNames.join(', ')}] and ${actionNames.length} action(s) [${actionNames.join(', ')}].`,

      // === RAW RESPONSE (for advanced debugging) ===
      _raw: raw,
    };
  }

  // =========================================================================
  // Flows -- WRITE Operations (Delegated User Token)
  // =========================================================================

  async createFlow(
    envId: string,
    displayName: string,
    definition: any,
    state: string = 'Stopped',
    connectionReferences?: any
  ): Promise<any> {
    definition = sanitizeDefinition(definition, displayName);

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

    fullDefinition = ensureRequiredParameters(fullDefinition, displayName);

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
      console.warn(`[Flow] \u26a0\ufe0f Definition for '${displayName}' has no triggers - flow may not execute`);
    }
    if (actionCount === 0) {
      console.warn(`[Flow] \u26a0\ufe0f Definition for '${displayName}' has no actions - flow will do nothing when triggered`);
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

    const newFlowId = result?.name || 'unknown';
    const flowDisplayName = result?.properties?.displayName || displayName;
    const flowState = result?.properties?.state || state;

    console.log(`[Flow] \u2705 Flow created: ${newFlowId} (${flowDisplayName})`);

    return {
      status: 'created',
      flowId: newFlowId,
      name: newFlowId,
      displayName: flowDisplayName,
      state: flowState,
      definition: fullDefinition,
      connectionReferences: connectionReferences || {},
      _source: 'flow-management-api',
      _authType: this.userAuth?.hasAuthenticatedUser() ? 'delegated' : 'service-principal',
      _note: 'Created via Flow Management API. If the create returned 200/201, the definition IS saved. ' +
        'Open the flow in Power Automate portal to authorize connections before enabling.',
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
      let def = sanitizeDefinition(updates.definition, `flow ${flowId}`);

      def = { ...def };
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
  // Flow Management -- LIFECYCLE (Flow Admin API)
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
  // TRIGGER FLOW (v3.0.2 -- fixed path)
  //
  // ROOT CAUSE (confirmed 2026-03-21):
  //   POST /scopes/admin/.../triggers/manual/run -> 404
  //   The /scopes/admin/ path does NOT support trigger execution.
  //
  // FIX: Use userFlowRequest with standard /environments/ path.
  //   manual trigger only works on flows with a Request (manual) trigger.
  //   Recurrence-triggered flows cannot be run via this endpoint.
  // =========================================================================

  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    console.log(`[Flow] Triggering flow ${flowId} via delegated user path (non-admin)`);
    console.log(`[Flow] Note: This only works for flows with a manual/Request trigger. ` +
      `Recurrence-triggered flows run on their schedule and cannot be triggered via API.`);

    return this.userFlowRequest(
      `/providers/Microsoft.ProcessSimple/environments/${envId}/flows/${flowId}/triggers/manual/run?api-version=${FLOW_API_VER}`,
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
