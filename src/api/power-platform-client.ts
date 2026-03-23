/**
 * PowerPlatformClient — HTTP client for Power Platform APIs
 *
 * v3.1.0: Added getConnectionDetails() and deleteConnection() methods
 *
 * v3.0.3: Dual-path GET for getFlowDetails with metadata fields:
 *   _fetchedVia, _definitionStatus, _definitionNote
 *
 * Architecture:
 *   - Service principal (AzureTokenManager) for admin read operations
 *   - Delegated user (UserAuthManager) for write operations + user-scoped reads
 *   - getFlowDetails: tries delegated first, falls back to admin
 *
 * @version 3.1.0
 */

import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';
import { UserAuthManager } from '../auth/user-auth-manager.js';

// =============================================================================
// Constants
// =============================================================================

const FLOW_API = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple';
const BAP_API = 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform';
const FLOW_API_VERSION = '2016-11-01';
const BAP_API_VERSION = '2023-06-01';

// =============================================================================
// PowerPlatformClient
// =============================================================================

export class PowerPlatformClient {
  private tokenManager: AzureTokenManager;
  private userAuthManager: UserAuthManager;

  constructor(tokenManager: AzureTokenManager, userAuthManager: UserAuthManager) {
    this.tokenManager = tokenManager;
    this.userAuthManager = userAuthManager;

    if (userAuthManager?.isConfigured()) {
      console.log('[PowerPlatformClient] Dual-token mode: service principal (read) + delegated user (write)');
    } else {
      console.log('[PowerPlatformClient] Service-principal only mode (no delegated auth configured)');
    }
  }

  // ===========================================================================
  // Internal Request Helpers
  // ===========================================================================

  /**
   * Admin request using service principal token (Flow Management API scope).
   * Used for: listFlows, getFlowDetails (fallback), enableDisable, delete, runs.
   */
  private async adminFlowRequest(path: string, method: string = 'GET', body?: any): Promise<any> {
    const token = await this.tokenManager.getToken('https://service.flow.microsoft.com/.default');
    const url = this.buildUrl(FLOW_API, path, FLOW_API_VERSION);

    const response = await axios({
      method,
      url,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      data: body,
      validateStatus: (s) => s < 500,
    });

    if (response.status >= 400) {
      const errMsg = response.data?.error?.message || response.data?.message || `HTTP ${response.status}`;
      throw new Error(`Admin API error (${response.status}): ${errMsg}`);
    }

    return response.data;
  }

  /**
   * User delegated request using per-user Device Code Flow token.
   * Used for: createFlow, updateFlow, triggerFlow, getFlowDetails (primary).
   */
  private async userFlowRequest(path: string, method: string = 'GET', body?: any, userId?: string): Promise<any> {
    const userToken = await this.userAuthManager.getAccessToken(userId);
    if (!userToken) {
      throw new Error(
        'No delegated user token available. ' +
        'Use pa-auth-start to begin Device Code Flow authentication before write operations.'
      );
    }

    const url = this.buildUrl(FLOW_API, path, FLOW_API_VERSION);

    const response = await axios({
      method,
      url,
      headers: {
        'Authorization': `Bearer ${userToken}`,
        'Content-Type': 'application/json',
      },
      data: body,
      validateStatus: (s) => s < 500,
    });

    if (response.status >= 400) {
      const errMsg = response.data?.error?.message || response.data?.message || `HTTP ${response.status}`;
      throw new Error(`User API error (${response.status}): ${errMsg}`);
    }

    return response.data;
  }

  /**
   * BAP admin request for environment operations.
   */
  private async bapRequest(path: string): Promise<any> {
    const token = await this.tokenManager.getToken('https://api.bap.microsoft.com/.default');
    const url = this.buildUrl(BAP_API, path, BAP_API_VERSION);

    const response = await axios.get(url, {
      headers: { 'Authorization': `Bearer ${token}` },
    });
    return response.data;
  }

  /**
   * Builds a full URL with api-version query parameter.
   */
  private buildUrl(base: string, path: string, apiVersion: string): string {
    const full = `${base}${path}`;
    const sep = full.includes('?') ? '&' : '?';
    return `${full}${sep}api-version=${apiVersion}`;
  }

  // ===========================================================================
  // Environment Operations
  // ===========================================================================

  async listEnvironments(): Promise<any> {
    return this.bapRequest('/environments');
  }

  // ===========================================================================
  // Flow Read Operations
  // ===========================================================================

  /**
   * Lists flows in an environment via admin scope.
   */
  async listFlows(envId: string, filter?: string, top?: number): Promise<any> {
    let path = `/scopes/admin/environments/${envId}/v2/flows`;
    const params: string[] = [];
    if (top) params.push(`$top=${top}`);
    if (filter && filter !== 'all') params.push(`$filter=${filter}`);
    if (params.length > 0) path += `?${params.join('&')}`;
    return this.adminFlowRequest(path);
  }

  /**
   * Gets detailed flow information with dual-path GET.
   *
   * v3.0.3: Tries delegated user endpoint first (returns full definition),
   * falls back to admin endpoint (may return empty definition for
   * delegated-created flows). Adds metadata fields:
   *   _fetchedVia: which API path was used
   *   _definitionStatus: POPULATED or EMPTY_OR_NOT_RETURNED
   *   _definitionNote: explanation for agents
   *   _raw: full API response (stripped by index.ts before sending to agent)
   */
  async getFlowDetails(envId: string, flowId: string): Promise<any> {
    let fetchedVia = '';
    let raw: any = null;
    let usedDelegated = false;

    // ---- Path 1: Try delegated user endpoint (returns full definition) ----
    if (this.userAuthManager.hasAuthenticatedUser()) {
      try {
        const userPath = `/environments/${envId}/flows/${flowId}`;
        raw = await this.userFlowRequest(userPath);
        const defaultUser = this.userAuthManager.getDefaultUserId();
        fetchedVia = `delegated (${defaultUser || 'user'})`;
        usedDelegated = true;
        console.log(`[getFlowDetails] Fetched via delegated user token (${defaultUser})`);
      } catch (e: any) {
        console.warn(`[getFlowDetails] Delegated GET failed (${e.message}), falling back to admin`);
      }
    }

    // ---- Path 2: Fall back to admin endpoint ----
    if (!raw) {
      const adminPath = `/scopes/admin/environments/${envId}/flows/${flowId}`;
      raw = await this.adminFlowRequest(adminPath);
      fetchedVia = 'admin (service-principal)';
      console.log('[getFlowDetails] Fetched via admin (service-principal)');
    }

    // ---- Extract definition details ----
    const definition = raw?.properties?.definition;
    const triggers = definition?.triggers ? Object.keys(definition.triggers) : [];
    const actions = definition?.actions ? Object.keys(definition.actions) : [];
    const hasDefinition = triggers.length > 0 || actions.length > 0;

    // ---- Build metadata fields for agent decision-making ----
    let definitionStatus: string;
    let definitionNote: string;

    if (hasDefinition) {
      definitionStatus = 'POPULATED';
      definitionNote = `Definition contains ${triggers.length} trigger(s) and ${actions.length} action(s).`;
    } else if (!usedDelegated) {
      // Empty via admin path — this is the known scope limitation
      definitionStatus = 'EMPTY_OR_NOT_RETURNED';
      definitionNote =
        'Definition not returned via admin endpoint. This is a KNOWN SCOPE LIMITATION — ' +
        'the admin API path does not return flow definitions for flows created via delegated auth. ' +
        'If the create/update returned 200, the definition IS saved. ' +
        'DO NOT delete this flow. ' +
        'To verify: re-authenticate with pa-auth-start and call pa-get-flow-details again, ' +
        'or ask the user to check the Power Automate portal designer.';
    } else {
      // Empty via delegated path — genuinely empty
      definitionStatus = 'EMPTY_OR_NOT_RETURNED';
      definitionNote =
        'Definition appears empty even via delegated user endpoint. ' +
        'The flow may genuinely have no definition, or it may not have been saved correctly. ' +
        'Consider recreating with a full definition (triggers + actions) in one call.';
    }

    return {
      id: raw?.name,
      displayName: raw?.properties?.displayName,
      state: raw?.properties?.state,
      createdTime: raw?.properties?.createdTime,
      lastModifiedTime: raw?.properties?.lastModifiedTime,
      triggers,
      actions,
      connectionReferences: raw?.properties?.connectionReferences,
      definition: definition || null,
      _fetchedVia: fetchedVia,
      _definitionStatus: definitionStatus,
      _definitionNote: definitionNote,
      _raw: raw,
    };
  }

  // ===========================================================================
  // Flow Write Operations (prefer delegated, fallback to admin)
  // ===========================================================================

  /**
   * Creates a new flow. Uses delegated user token (preferred) or admin fallback.
   * Sends shell + definition in ONE API call.
   */
  async createFlow(
    envId: string,
    displayName: string,
    definition: any,
    state?: string,
    connectionReferences?: any
  ): Promise<any> {
    const body: any = {
      properties: {
        displayName,
        definition: definition || {},
        state: state || 'Stopped',
      },
    };
    if (connectionReferences) {
      body.properties.connectionReferences = connectionReferences;
    }

    // Prefer delegated user token for writes
    if (this.userAuthManager.hasAuthenticatedUser()) {
      try {
        const path = `/environments/${envId}/flows`;
        const result = await this.userFlowRequest(path, 'POST', body);
        console.log(`[createFlow] Flow created via delegated user token: ${result?.name}`);
        return {
          status: 'created',
          flowId: result?.name,
          displayName: result?.properties?.displayName || displayName,
          state: result?.properties?.state || state || 'Stopped',
          _source: 'user-endpoint',
          _authType: 'delegated',
          _idMapping: result?.name ? `Flow ID: ${result.name}` : undefined,
        };
      } catch (e: any) {
        console.error(`[createFlow] Delegated POST failed: ${e.message}`);
        throw new Error(
          `Flow creation failed: ${e.message}. ` +
          'Ensure you are authenticated (pa-auth-start) and the token has not expired.'
        );
      }
    }

    // No delegated token — attempt admin fallback
    try {
      const path = `/scopes/admin/environments/${envId}/flows`;
      const result = await this.adminFlowRequest(path, 'POST', body);
      console.log(`[createFlow] Flow created via admin (service-principal): ${result?.name}`);
      return {
        status: 'created',
        flowId: result?.name,
        displayName: result?.properties?.displayName || displayName,
        state: result?.properties?.state || state || 'Stopped',
        _source: 'admin-endpoint',
        _authType: 'service-principal',
        _idMapping: result?.name ? `Flow ID: ${result.name}` : undefined,
      };
    } catch (e: any) {
      throw new Error(`Flow creation failed: ${e.message}`);
    }
  }

  /**
   * Updates an existing flow. Uses delegated user token (preferred) or admin fallback.
   */
  async updateFlow(envId: string, flowId: string, updates: any): Promise<any> {
    const body: any = { properties: {} };
    if (updates.displayName) body.properties.displayName = updates.displayName;
    if (updates.definition) body.properties.definition = updates.definition;
    if (updates.state) body.properties.state = updates.state;
    if (updates.connectionReferences) body.properties.connectionReferences = updates.connectionReferences;

    // Prefer delegated user token for writes
    if (this.userAuthManager.hasAuthenticatedUser()) {
      try {
        const path = `/environments/${envId}/flows/${flowId}`;
        await this.userFlowRequest(path, 'PATCH', body);
        console.log(`[updateFlow] Flow ${flowId} updated via delegated user token`);
        return {
          status: 'updated',
          flowId,
          updatedProperties: Object.keys(updates),
          _source: 'user-endpoint',
          _authType: 'delegated',
        };
      } catch (e: any) {
        console.error(`[updateFlow] Delegated PATCH failed: ${e.message}`);
        throw new Error(
          `Flow update failed: ${e.message}. ` +
          'Ensure you are authenticated (pa-auth-start) and the token has not expired.'
        );
      }
    }

    // Fallback to admin
    try {
      const path = `/scopes/admin/environments/${envId}/flows/${flowId}`;
      await this.adminFlowRequest(path, 'PATCH', body);
      console.log(`[updateFlow] Flow ${flowId} updated via admin (service-principal)`);
      return {
        status: 'updated',
        flowId,
        updatedProperties: Object.keys(updates),
        _source: 'admin-endpoint',
        _authType: 'service-principal',
      };
    } catch (e: any) {
      throw new Error(`Flow update failed: ${e.message}`);
    }
  }

  // ===========================================================================
  // Flow Lifecycle Operations
  // ===========================================================================

  /**
   * Enables (start) or disables (stop) a flow.
   */
  async enableDisableFlow(envId: string, flowId: string, action: string): Promise<void> {
    const path = `/scopes/admin/environments/${envId}/flows/${flowId}/${action}`;
    await this.adminFlowRequest(path, 'POST');
    console.log(`[enableDisableFlow] Flow ${flowId}: ${action}`);
  }

  /**
   * Permanently deletes a flow.
   */
  async deleteFlow(envId: string, flowId: string): Promise<void> {
    // Try delegated first (user may own the flow)
    if (this.userAuthManager.hasAuthenticatedUser()) {
      try {
        const path = `/environments/${envId}/flows/${flowId}`;
        await this.userFlowRequest(path, 'DELETE');
        console.log(`[deleteFlow] Flow ${flowId} deleted via delegated user token`);
        return;
      } catch (e: any) {
        console.warn(`[deleteFlow] Delegated DELETE failed (${e.message}), trying admin`);
      }
    }

    const path = `/scopes/admin/environments/${envId}/flows/${flowId}`;
    await this.adminFlowRequest(path, 'DELETE');
    console.log(`[deleteFlow] Flow ${flowId} deleted via admin (service-principal)`);
  }

  /**
   * Manually triggers a flow. Requires delegated auth for user-scoped flows.
   * Corrected trigger path (v3.0.2): /triggers/manual/run
   */
  async triggerFlow(envId: string, flowId: string, body?: any): Promise<any> {
    // Prefer delegated endpoint for trigger
    if (this.userAuthManager.hasAuthenticatedUser()) {
      try {
        const path = `/environments/${envId}/flows/${flowId}/triggers/manual/run`;
        const result = await this.userFlowRequest(path, 'POST', body || {});
        console.log(`[triggerFlow] Flow ${flowId} triggered via delegated user token`);
        return result;
      } catch (e: any) {
        console.warn(`[triggerFlow] Delegated trigger failed (${e.message}), trying admin`);
      }
    }

    // Fallback to admin trigger path
    const path = `/scopes/admin/environments/${envId}/flows/${flowId}/triggers/manual/run`;
    const result = await this.adminFlowRequest(path, 'POST', body || {});
    console.log(`[triggerFlow] Flow ${flowId} triggered via admin (service-principal)`);
    return result;
  }

  // ===========================================================================
  // Flow Run Operations
  // ===========================================================================

  /**
   * Gets run history for a flow.
   */
  async getRunHistory(envId: string, flowId: string, top?: number): Promise<any> {
    let path = `/scopes/admin/environments/${envId}/flows/${flowId}/runs`;
    if (top) path += `?$top=${top}`;
    return this.adminFlowRequest(path);
  }

  /**
   * Gets details for a specific flow run.
   */
  async getRunDetails(envId: string, flowId: string, runId: string): Promise<any> {
    const path = `/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}`;
    return this.adminFlowRequest(path);
  }

  /**
   * Cancels a running flow execution.
   */
  async cancelRun(envId: string, flowId: string, runId: string): Promise<void> {
    const path = `/scopes/admin/environments/${envId}/flows/${flowId}/runs/${runId}/cancel`;
    await this.adminFlowRequest(path, 'POST');
    console.log(`[cancelRun] Run ${runId} cancelled for flow ${flowId}`);
  }

  // ===========================================================================
  // Connection Operations
  // ===========================================================================

  /**
   * Lists connections in an environment.
   */
  async listConnections(envId: string): Promise<any> {
    const path = `/scopes/admin/environments/${envId}/connections`;
    return this.adminFlowRequest(path);
  }

  /**
   * Gets details of a specific connection.
   * v3.1.0: New method — wired to pa-get-connection tool.
   */
  async getConnectionDetails(envId: string, connectionName: string): Promise<any> {
    const path = `/scopes/admin/environments/${envId}/connections/${connectionName}`;
    return this.adminFlowRequest(path);
  }

  /**
   * Deletes a connection permanently.
   * v3.1.0: New method — wired to pa-delete-connection tool.
   *
   * WARNING: Deleting a connection may break flows that depend on it.
   */
  async deleteConnection(envId: string, connectionName: string): Promise<void> {
    const path = `/scopes/admin/environments/${envId}/connections/${connectionName}`;
    await this.adminFlowRequest(path, 'DELETE');
    console.log(`[deleteConnection] Connection ${connectionName} deleted`);
  }
}
