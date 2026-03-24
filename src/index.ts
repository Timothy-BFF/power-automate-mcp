import 'dotenv/config';
import express, { Request, Response } from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { z } from 'zod';
import axios from 'axios';
import { AzureTokenManager } from './auth/azure-token-manager.js';
import { PowerPlatformClient } from './api/power-platform-client.js';
import { SolutionClient } from './clients/solution-client.js';
import { resolveEnvironmentId } from './config/environment-resolver.js';
import { ToolResult, ToolDefinition } from './types.js';
import { UserAuthManager } from './auth/user-auth-manager.js';
import { TOOL_DESCRIPTIONS } from './tools/tool-descriptions.js';

// =============================================================================
// Configuration
// =============================================================================
const PORT = parseInt(process.env.PORT || '8080', 10);
const VERSION = '3.3.0';

// =============================================================================
// Utility
// =============================================================================
function generateUUID(): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

// =============================================================================
// Core Services
// =============================================================================
const tokenManager = new AzureTokenManager();
const userAuthManager = new UserAuthManager();
const client = new PowerPlatformClient(tokenManager, userAuthManager);
const solutionClient = new SolutionClient(tokenManager);

// Pre-warm tokens (BAP + Flow + PowerApps + Dataverse scopes)
(async () => {
  try {
    await tokenManager.getToken('https://api.bap.microsoft.com/.default');
    console.log('[Init] BAP admin token acquired (api.bap.microsoft.com)');
  } catch (e: any) { console.warn('[Init] BAP token pre-warm failed:', e.message); }
  try {
    await tokenManager.getToken('https://service.flow.microsoft.com/.default');
    console.log('[Init] Flow token acquired (service.flow.microsoft.com)');
  } catch (e: any) { console.warn('[Init] Flow token pre-warm failed:', e.message); }
  try {
    await tokenManager.getToken('https://service.powerapps.com/.default');
    console.log('[Init] PowerApps token acquired (service.powerapps.com)');
  } catch (e: any) { console.warn('[Init] PowerApps token pre-warm failed:', e.message); }
  // Dataverse scope (if configured)
  if (solutionClient.isConfigured()) {
    try {
      await tokenManager.getToken(solutionClient.getScope());
      console.log(`[Init] Dataverse token acquired (${process.env.DATAVERSE_URL})`);
    } catch (e: any) { console.warn('[Init] Dataverse token pre-warm failed:', e.message); }
  }
})();

// =============================================================================
// Tool Result Helpers
// =============================================================================
function ok(data: any): ToolResult {
  return { content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }] };
}
function fail(msg: string): ToolResult {
  return { content: [{ type: 'text' as const, text: JSON.stringify({ error: msg }) }], isError: true };
}

// =============================================================================
// Tool Definitions (shared between SSE + REST transports)
//
// v3.3.0: Fixed pa-list-connections (PowerApps host + admin token),
//         Fixed pa-create-connection (multi-scope refresh token exchange +
//         correct PUT endpoint with connector in path + env as filter).
//         Added pa-create-solution, pa-delete-solution, pa-create-connection.
//         Unified tool registration — all 26 tools through one loop.
//         Total: 26 tools.
// v3.2.0: Added 5 Dataverse Solutions tools.
// v3.1.0: Added pa-get-connection and pa-delete-connection tools.
// v3.0.3: All descriptions sourced from TOOL_DESCRIPTIONS.
// =============================================================================
const toolDefs: ToolDefinition[] = [
  // ---- List Environments ----
  {
    name: 'pa-list-environments',
    description: TOOL_DESCRIPTIONS['pa-list-environments'],
    inputSchema: { type: 'object', properties: {} },
    handler: async () => {
      try {
        const r = await client.listEnvironments();
        const envs = (r.value || []).map((e: any) => ({
          id: e.name, displayName: e.properties?.displayName, location: e.location,
          sku: e.properties?.environmentSku, state: e.properties?.states?.management?.id,
        }));
        return ok({ count: envs.length, environments: envs });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- List Flows ----
  {
    name: 'pa-list-flows',
    description: TOOL_DESCRIPTIONS['pa-list-flows'],
    inputSchema: {
      type: 'object',
      properties: {
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
        filter: { type: 'string', description: 'Filter: personal, shared, or all.' },
        top: { type: 'number', description: 'Maximum number of flows to return.' },
      },
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.listFlows(envId, p.filter, p.top ? Number(p.top) : undefined);
        const flows = (r.value || []).map((f: any) => ({
          id: f.name, displayName: f.properties?.displayName, state: f.properties?.state,
          createdTime: f.properties?.createdTime, lastModifiedTime: f.properties?.lastModifiedTime,
        }));
        return ok({ count: flows.length, flows });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Get Flow Details ----
  {
    name: 'pa-get-flow-details',
    description: TOOL_DESCRIPTIONS['pa-get-flow-details'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to retrieve details for.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.getFlowDetails(envId, p.flowId);
        if (result && result._raw) {
          const { _raw, ...clean } = result;
          return ok(clean);
        }
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Create Flow ----
  {
    name: 'pa-create-flow',
    description: TOOL_DESCRIPTIONS['pa-create-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        displayName: { type: 'string', description: 'Flow display name.' },
        definition: { type: 'object', description: 'Complete workflow definition JSON with triggers and actions.' },
        state: { type: 'string', description: 'Initial state: Started or Stopped (default: Stopped).' },
        connectionReferences: { type: 'object', description: 'Connection references for connectors.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['displayName', 'definition'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.createFlow(envId, p.displayName, p.definition, p.state, p.connectionReferences);
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Update Flow ----
  {
    name: 'pa-update-flow',
    description: TOOL_DESCRIPTIONS['pa-update-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to update.' },
        displayName: { type: 'string', description: 'New display name.' },
        definition: { type: 'object', description: 'Updated workflow definition JSON.' },
        state: { type: 'string', description: 'New state: Started or Stopped.' },
        connectionReferences: { type: 'object', description: 'Updated connection references.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const updates: any = {};
        if (p.displayName) updates.displayName = p.displayName;
        if (p.definition) updates.definition = p.definition;
        if (p.state) updates.state = p.state;
        if (p.connectionReferences) updates.connectionReferences = p.connectionReferences;
        const result = await client.updateFlow(envId, p.flowId, updates);
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Enable Flow ----
  {
    name: 'pa-enable-flow',
    description: TOOL_DESCRIPTIONS['pa-enable-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to enable.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        await client.enableDisableFlow(envId, p.flowId, 'start');
        return ok({ status: 'enabled', flowId: p.flowId, message: `Flow ${p.flowId} has been enabled.` });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Disable Flow ----
  {
    name: 'pa-disable-flow',
    description: TOOL_DESCRIPTIONS['pa-disable-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to disable.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        await client.enableDisableFlow(envId, p.flowId, 'stop');
        return ok({ status: 'disabled', flowId: p.flowId, message: `Flow ${p.flowId} has been disabled.` });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Delete Flow ----
  {
    name: 'pa-delete-flow',
    description: TOOL_DESCRIPTIONS['pa-delete-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to delete.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        await client.deleteFlow(envId, p.flowId);
        return ok({ status: 'deleted', flowId: p.flowId, message: `Flow ${p.flowId} has been permanently deleted.` });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Trigger Flow ----
  {
    name: 'pa-trigger-flow',
    description: TOOL_DESCRIPTIONS['pa-trigger-flow'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to trigger.' },
        body: { type: 'object', description: 'Request body for the trigger.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.triggerFlow(envId, p.flowId, p.body);
        return ok({ status: 'triggered', flowId: p.flowId, result });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Get Run History ----
  {
    name: 'pa-get-run-history',
    description: TOOL_DESCRIPTIONS['pa-get-run-history'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID to get run history for.' },
        top: { type: 'number', description: 'Maximum number of runs to return.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.getRunHistory(envId, p.flowId, p.top ? Number(p.top) : undefined);
        const runs = (result?.value || []).map((r: any) => ({
          runId: r.name, status: r.properties?.status,
          startTime: r.properties?.startTime, endTime: r.properties?.endTime,
          trigger: r.properties?.trigger?.name,
        }));
        return ok({ flowId: p.flowId, count: runs.length, runs });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Get Run Details ----
  {
    name: 'pa-get-run-details',
    description: TOOL_DESCRIPTIONS['pa-get-run-details'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID.' },
        runId: { type: 'string', description: 'Run ID to get details for.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId', 'runId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.getRunDetails(envId, p.flowId, p.runId);
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Cancel Run ----
  {
    name: 'pa-cancel-run',
    description: TOOL_DESCRIPTIONS['pa-cancel-run'],
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'Flow ID.' },
        runId: { type: 'string', description: 'Run ID to cancel.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['flowId', 'runId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        await client.cancelRun(envId, p.flowId, p.runId);
        return ok({ status: 'cancelled', flowId: p.flowId, runId: p.runId });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // =========================================================================
  // CONNECTION TOOLS
  // v3.3.0: Fixed pa-list-connections (PowerApps host + admin token),
  //         Fixed pa-create-connection (multi-scope token + correct PUT endpoint)
  // =========================================================================
  // ---- List Connections (v3.3.0 FIX: PowerApps host + PowerApps token) ----
  {
    name: 'pa-list-connections',
    description: TOOL_DESCRIPTIONS['pa-list-connections'],
    inputSchema: {
      type: 'object',
      properties: {
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);

        // v3.3.0 FIX: Connections API lives under PowerApps, NOT Flow.
        const token = await tokenManager.getToken('https://service.powerapps.com/.default');
        const url = `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${envId}/connections?api-version=2016-11-01`;

        console.log(`[ListConnections] GET ${url} (PowerApps admin token)`);

        const response = await axios.get(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
        });

        const connections = (response.data?.value || []).map((c: any) => ({
          connectionId: c.name,
          displayName: c.properties?.displayName,
          apiId: c.properties?.apiId,
          status: c.properties?.statuses?.[0]?.status,
          createdTime: c.properties?.createdTime,
        }));
        return ok({
          count: connections.length,
          connections,
          _apiHost: 'api.powerapps.com',
          _tokenScope: 'service.powerapps.com',
        });
      } catch (e: any) {
        const status = e.response?.status;
        const detail = e.response?.data?.error?.message || e.message;
        return fail(`ListConnections failed (${status || 'unknown'}): ${detail}`);
      }
    },
  },
  // ---- Get Connection Details ----
  {
    name: 'pa-get-connection',
    description: TOOL_DESCRIPTIONS['pa-get-connection'],
    inputSchema: {
      type: 'object',
      properties: {
        connectionId: { type: 'string', description: 'Connection ID (name field from pa-list-connections).' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['connectionId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const result = await client.getConnectionDetails(envId, p.connectionId);
        const c = result;
        return ok({
          connectionId: c.name,
          displayName: c.properties?.displayName,
          apiId: c.properties?.apiId,
          connectorDisplayName: c.properties?.apiId?.split('/').pop(),
          status: c.properties?.statuses?.[0]?.status,
          createdTime: c.properties?.createdTime,
          createdBy: c.properties?.createdBy,
          statuses: c.properties?.statuses,
          connectionParameters: c.properties?.connectionParametersSet || null,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Delete Connection ----
  {
    name: 'pa-delete-connection',
    description: TOOL_DESCRIPTIONS['pa-delete-connection'],
    inputSchema: {
      type: 'object',
      properties: {
        connectionId: { type: 'string', description: 'Connection ID (name field) to delete.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      },
      required: ['connectionId'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        await client.deleteConnection(envId, p.connectionId);
        return ok({
          status: 'deleted',
          connectionId: p.connectionId,
          message: `Connection ${p.connectionId} has been permanently deleted.`,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // =========================================================================
  // Create Connection (v3.3.0 FIX round 3)
  //
  // Previous attempts:
  //   Round 1 (930cda56): POST /environments/{env}/connections → 403 (wrong token scope)
  //   Round 2 (7e1aef7b): POST /environments/{env}/connections + PowerApps token → 404 (wrong URL)
  //
  // Fix: The PowerApps connection creation API uses a completely different
  // URL pattern than listing. The connector goes IN the URL path and the
  // environment is passed as a query filter:
  //
  //   PUT /providers/Microsoft.PowerApps/apis/{connectorName}/connections/{newId}
  //       ?api-version=2016-11-01&$filter=environment eq '{envId}'
  //
  // For OAuth connectors (shared_office365, etc.), the API creates the
  // connection shell but the user must authorize it in the Power Automate
  // portal. The response may include a consentLink or an Error status.
  // =========================================================================
  {
    name: 'pa-create-connection',
    description: TOOL_DESCRIPTIONS['pa-create-connection'],
    inputSchema: {
      type: 'object',
      properties: {
        connectorId: { type: 'string', description: 'Connector API name (e.g., shared_office365, shared_sql, shared_sharepointonline) or full API ID path.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID.' },
        user_id: { type: 'string', description: 'Email of the authenticated user who will own this connection.' },
        connectionParameters: { type: 'object', description: 'Optional connection parameters (varies by connector, e.g., server/database for SQL).' },
      },
      required: ['connectorId', 'user_id'],
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        if (!p.user_id) return fail('user_id is required for creating connections. Use pa-auth-start first.');

        // Acquire PowerApps-scoped delegated token via refresh token exchange
        const userToken = await userAuthManager.getAccessTokenForScope(
          p.user_id,
          'https://service.powerapps.com/.default'
        );
        if (!userToken) {
          return fail(
            `Could not acquire PowerApps delegated token for ${p.user_id}. ` +
            'Ensure the user is authenticated (pa-auth-start) and the refresh token is valid. ' +
            'If this persists, admin consent for PowerApps Service may be required in Azure AD.'
          );
        }

        // Extract short connector name (e.g., shared_office365)
        const connectorName = p.connectorId.startsWith('/providers/')
          ? p.connectorId.split('/').pop()!
          : p.connectorId;

        // Generate unique connection name: shared-office365-{uuid}
        const connectionName = `${connectorName.replace(/_/g, '-')}-${generateUUID()}`;

        // Correct PowerApps connection creation endpoint:
        //   PUT /apis/{connector}/connections/{name}?$filter=environment eq '{env}'
        // NOT: POST /environments/{env}/connections (that's only for listing)
        const url = `https://api.powerapps.com/providers/Microsoft.PowerApps/apis/${connectorName}/connections/${connectionName}`;

        console.log(`[CreateConnection] PUT ${url} (user: ${p.user_id}, env: ${envId}, token: powerapps-delegated)`);

        const response = await axios.put(url, {
          properties: {
            environment: {
              id: `/providers/Microsoft.PowerApps/environments/${envId}`,
              name: envId,
            },
            connectionParameters: p.connectionParameters || {},
          },
        }, {
          headers: {
            Authorization: `Bearer ${userToken}`,
            'Content-Type': 'application/json',
          },
          params: {
            'api-version': '2016-11-01',
            '$filter': `environment eq '${envId}'`,
          },
        });

        const conn = response.data;
        const connStatus = conn?.properties?.statuses?.[0]?.status || 'Unknown';

        // Check if the connector requires interactive consent (OAuth connectors)
        const needsConsent = connStatus === 'Error' || connStatus === 'Unauthenticated';
        const consentLink = conn?.properties?.connectionParameters?.token?.oAuthSettings?.redirectUrl || null;

        return ok({
          status: 'created',
          connectionId: conn?.name || connectionName,
          displayName: conn?.properties?.displayName || connectionName,
          connectorName,
          connectionStatus: connStatus,
          needsConsent,
          consentLink,
          _tokenScope: 'service.powerapps.com (delegated)',
          _method: 'PUT',
          _urlPattern: 'apis/{connector}/connections/{name}?$filter=environment',
          message: needsConsent
            ? `Connection shell created for ${connectorName}. This OAuth connector requires interactive authorization. ` +
              `Go to make.powerautomate.com → Data → Connections → find "${connectionName}" → Authorize. ` +
              (consentLink ? `Or visit: ${consentLink}` : '')
            : `Connection created successfully. Status: ${connStatus}.`,
        });
      } catch (e: any) {
        const status = e.response?.status;
        const errCode = e.response?.data?.error?.code || '';
        const detail = e.response?.data?.error?.message || e.message;

        if (status === 404) {
          return fail(
            `CreateConnection 404: Connector '${p.connectorId}' not found at the PowerApps API. ` +
            `This connector may not support API-based creation, or requires interactive consent. ` +
            `For OAuth connectors (Office 365, Outlook, SharePoint), create manually at make.powerautomate.com → Data → Connections. ` +
            `Detail: ${detail}`
          );
        }
        if (status === 403) {
          return fail(
            `CreateConnection 403 Forbidden: ${detail}. ` +
            'Admin consent for PowerApps Service may be required in Azure AD.'
          );
        }
        if (status === 409) {
          return fail(
            `CreateConnection 409 Conflict: A connection with this name already exists. ` +
            `Detail: ${detail}`
          );
        }
        return fail(`CreateConnection failed (${status || 'unknown'}): ${errCode} — ${detail}`);
      }
    },
  },
  // =========================================================================
  // DATAVERSE SOLUTIONS TOOLS (v3.2.0 + v3.3.0 additions)
  // =========================================================================
  // ---- List Solutions ----
  {
    name: 'pa-list-solutions',
    description: TOOL_DESCRIPTIONS['pa-list-solutions'],
    inputSchema: {
      type: 'object',
      properties: {
        includeManaged: { type: 'boolean', description: 'Include managed solutions (default: false, shows unmanaged only).' },
      },
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const solutions = await solutionClient.listSolutions(p.includeManaged === true);
        return ok({ count: solutions.length, solutions });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Get Solution Details ----
  {
    name: 'pa-get-solution',
    description: TOOL_DESCRIPTIONS['pa-get-solution'],
    inputSchema: {
      type: 'object',
      properties: {
        solutionId: { type: 'string', description: 'Solution GUID or unique name.' },
      },
      required: ['solutionId'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const solution = await solutionClient.getSolution(p.solutionId);
        return ok(solution);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- List Solution Components ----
  {
    name: 'pa-list-solution-components',
    description: TOOL_DESCRIPTIONS['pa-list-solution-components'],
    inputSchema: {
      type: 'object',
      properties: {
        solutionId: { type: 'string', description: 'Solution GUID (from pa-list-solutions or pa-get-solution).' },
      },
      required: ['solutionId'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const components = await solutionClient.listSolutionComponents(p.solutionId);
        const byType: Record<string, number> = {};
        for (const c of components) {
          byType[c.componentTypeName] = (byType[c.componentTypeName] || 0) + 1;
        }
        return ok({ count: components.length, componentsByType: byType, components });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Export Solution ----
  {
    name: 'pa-export-solution',
    description: TOOL_DESCRIPTIONS['pa-export-solution'],
    inputSchema: {
      type: 'object',
      properties: {
        solutionUniqueName: { type: 'string', description: 'Solution unique name (from pa-list-solutions).' },
        managed: { type: 'boolean', description: 'Export as managed solution (default: false).' },
      },
      required: ['solutionUniqueName'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const result = await solutionClient.exportSolution(
          p.solutionUniqueName,
          p.managed === true
        );
        return ok({
          status: 'exported',
          fileName: result.fileName,
          sizeBytes: result.sizeBytes,
          sizeHuman: result.sizeBytes < 1024
            ? `${result.sizeBytes} bytes`
            : result.sizeBytes < 1048576
              ? `${Math.round(result.sizeBytes / 1024)}KB`
              : `${(result.sizeBytes / 1048576).toFixed(1)}MB`,
          managed: p.managed === true,
          base64Content: result.base64Content,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Add Solution Component ----
  {
    name: 'pa-add-solution-component',
    description: TOOL_DESCRIPTIONS['pa-add-solution-component'],
    inputSchema: {
      type: 'object',
      properties: {
        solutionUniqueName: { type: 'string', description: 'Solution unique name to add the component to.' },
        componentId: { type: 'string', description: 'GUID of the component (e.g., flow ID).' },
        componentType: { type: 'number', description: 'Component type code (default: 29 = Cloud Flow).' },
        addRequiredComponents: { type: 'boolean', description: 'Auto-include dependencies (default: false).' },
      },
      required: ['solutionUniqueName', 'componentId'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const result = await solutionClient.addSolutionComponent(
          p.solutionUniqueName,
          p.componentId,
          p.componentType != null ? Number(p.componentType) : 29,
          p.addRequiredComponents === true
        );
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Create Solution (v3.3.0 NEW) ----
  {
    name: 'pa-create-solution',
    description: TOOL_DESCRIPTIONS['pa-create-solution'],
    inputSchema: {
      type: 'object',
      properties: {
        uniqueName: { type: 'string', description: 'Unique name for the solution (no spaces, e.g., MyNewSolution).' },
        friendlyName: { type: 'string', description: 'Display name for the solution.' },
        publisherId: { type: 'string', description: 'GUID of the publisher. Use publisherId from pa-get-solution on an existing solution.' },
        version: { type: 'string', description: 'Version number (default: 1.0.0.0).' },
        description: { type: 'string', description: 'Optional description.' },
      },
      required: ['uniqueName', 'friendlyName', 'publisherId'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const result = await solutionClient.createSolution(
          p.uniqueName,
          p.friendlyName,
          p.publisherId,
          p.version || '1.0.0.0',
          p.description || ''
        );
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Delete Solution (v3.3.0 NEW) ----
  {
    name: 'pa-delete-solution',
    description: TOOL_DESCRIPTIONS['pa-delete-solution'],
    inputSchema: {
      type: 'object',
      properties: {
        solutionId: { type: 'string', description: 'Solution GUID to delete (solutionId from pa-get-solution).' },
      },
      required: ['solutionId'],
    },
    handler: async (p: any) => {
      try {
        if (!solutionClient.isConfigured()) return fail('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        const result = await solutionClient.deleteSolution(p.solutionId);
        return ok(result);
      } catch (e: any) { return fail(e.message); }
    },
  },
  // =========================================================================
  // AUTH TOOLS (Device Code Flow)
  // v3.3.0: Registered via unified loop (no separate registerAuthTools call)
  // =========================================================================
  // ---- Auth: Start Device Code Flow ----
  {
    name: 'pa-auth-start',
    description: TOOL_DESCRIPTIONS['pa-auth-start'],
    inputSchema: {
      type: 'object',
      properties: {
        user_id: { type: 'string', description: 'User email (e.g., jose@company.com).' },
      },
      required: ['user_id'],
    },
    handler: async (p: any) => {
      try {
        if (!p.user_id) return fail('user_id is required. Provide the user email address.');
        const result = await userAuthManager.startAuth(p.user_id);
        return ok({
          status: 'device_code_issued',
          user_code: result.userCode,
          verification_uri: result.verificationUri,
          message: result.message,
          expires_in: result.expiresIn,
          _note: 'After user signs in, call pa-auth-poll to complete authentication.',
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Auth: Poll ----
  {
    name: 'pa-auth-poll',
    description: TOOL_DESCRIPTIONS['pa-auth-poll'],
    inputSchema: {
      type: 'object',
      properties: {
        user_id: { type: 'string', description: 'User email to poll for.' },
      },
    },
    handler: async (p: any) => {
      try {
        const targetUser = p.user_id || userAuthManager.getDefaultUserId();
        if (!targetUser) return fail('No user_id provided and no pending auth found. Call pa-auth-start first.');
        const result = await userAuthManager.pollAuth(targetUser);
        if (result.status === 'authenticated') {
          return ok({
            status: 'authenticated',
            user_id: result.userId || targetUser,
            message: result.message,
            _note: 'You can now use pa-create-flow, pa-update-flow, pa-list-connections, and pa-create-connection.',
          });
        }
        if (result.status === 'pending') {
          return ok({ status: 'pending', message: result.message });
        }
        return ok({ status: result.status, message: result.message });
      } catch (e: any) { return fail(e.message); }
    },
  },
  // ---- Auth: Status ----
  {
    name: 'pa-auth-status',
    description: TOOL_DESCRIPTIONS['pa-auth-status'],
    inputSchema: {
      type: 'object',
      properties: {
        user_id: { type: 'string', description: 'User email to check.' },
      },
    },
    handler: async (p: any) => {
      try {
        const status = userAuthManager.getAuthStatus(p.user_id);
        return ok({
          authenticated: status.authenticated,
          user_id: status.userId,
          message: status.message,
          pending: status.pending,
          authenticatedUsers: status.authenticatedUsers,
          expiresIn: status.expiresIn,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
];

// =============================================================================
// Handler Map (for REST JSON-RPC transport)
// =============================================================================
const toolHandlers = new Map<string, (params: any) => Promise<ToolResult>>();
for (const def of toolDefs) {
  toolHandlers.set(def.name, def.handler);
}

// =============================================================================
// MCP Server (SSE transport)
// =============================================================================
const mcpServer = new McpServer({
  name: 'power-automate-mcp',
  version: VERSION,
});

// Convert JSON Schema properties to zod for McpServer.tool() registration
function toZodProps(inputSchema: any): Record<string, any> {
  const props: Record<string, any> = {};
  const required = inputSchema?.required || [];
  for (const [key, val] of Object.entries((inputSchema?.properties || {}) as Record<string, any>)) {
    const isReq = required.includes(key);
    let s;
    if (val.type === 'number') {
      s = z.number().describe(val.description || key);
    } else if (val.type === 'boolean') {
      s = z.boolean().describe(val.description || key);
    } else if (val.type === 'object' || val.type === 'array') {
      s = z.any().describe(val.description || key);
    } else {
      s = z.string().describe(val.description || key);
    }
    props[key] = isReq ? s : s.optional();
  }
  return props;
}

// =============================================================================
// UNIFIED TOOL REGISTRATION (v3.3.0 fix)
// =============================================================================
for (const def of toolDefs) {
  const zodProps = toZodProps(def.inputSchema);
  mcpServer.tool(def.name, def.description, zodProps, async (params: any) => def.handler(params));
}

console.log(`[Init] MCP tools registered (SSE): ${toolDefs.length}`);

// =============================================================================
// Express App
// =============================================================================
const app = express();
const jsonParser = express.json();
const sseTransports = new Map<string, SSEServerTransport>();

// Health endpoint
app.get('/health', (_req: Request, res: Response) => {
  res.json({
    status: 'ok',
    version: VERSION,
    tools: toolDefs.length,
    transport: ['sse', 'rest'],
    auth: userAuthManager.isConfigured() ? 'dual-token' : 'service-principal-only',
    dataverse: solutionClient.isConfigured() ? 'configured' : 'not-configured',
  });
});

// --- SSE Transport ---
app.get('/sse', async (_req: Request, res: Response) => {
  try {
    await mcpServer.close();
    console.log('[SSE] Previous transport closed (reconnect detected)');
  } catch (_) {
    // No existing connection - first connect.
  }

  const transport = new SSEServerTransport('/messages', res);
  sseTransports.set(transport.sessionId, transport);
  res.on('close', () => {
    console.log(`[SSE] Connection closed: ${transport.sessionId}`);
    sseTransports.delete(transport.sessionId);
  });
  await mcpServer.connect(transport);
});

app.post('/messages', async (req: Request, res: Response) => {
  const sessionId = req.query.sessionId as string;
  const transport = sseTransports.get(sessionId);
  if (!transport) { res.status(404).json({ error: 'Session not found', sessionId }); return; }
  await transport.handlePostMessage(req, res);
});

// --- REST JSON-RPC Handler ---
async function processJsonRpcRequest(request: any): Promise<any> {
  const { jsonrpc, id, method, params } = request || {};

  if (method && toolHandlers.has(method)) {
    try {
      const result = await toolHandlers.get(method)!(params || {});
      return { jsonrpc: '2.0', id: id || null, result };
    } catch (e: any) {
      return { jsonrpc: '2.0', id: id || null, error: { code: -32603, message: e.message } };
    }
  }

  try {
    switch (method) {
      case 'initialize':
        return {
          jsonrpc: '2.0', id,
          result: {
            protocolVersion: '2024-11-05',
            capabilities: { tools: { listChanged: false } },
            serverInfo: { name: 'power-automate-mcp', version: VERSION },
          },
        };
      case 'notifications/initialized':
      case 'initialized':
        return { jsonrpc: '2.0', id, result: {} };
      case 'tools/list':
        return {
          jsonrpc: '2.0', id,
          result: { tools: toolDefs.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })) },
        };
      case 'tools/call': {
        const toolName = params?.name;
        const toolArgs = params?.arguments || {};
        const handler = toolHandlers.get(toolName);
        if (!handler) {
          return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: JSON.stringify({ error: `Unknown tool: ${toolName}` }) }], isError: true } };
        }
        const result = await handler(toolArgs);
        return { jsonrpc: '2.0', id, result };
      }
      case 'ping':
        return { jsonrpc: '2.0', id, result: {} };
      default:
        return { jsonrpc: '2.0', id: id || null, error: { code: -32601, message: `Method not found: ${method}` } };
    }
  } catch (err: any) {
    console.error(`[JSON-RPC] Error in ${method}:`, err.message);
    return { jsonrpc: '2.0', id: id || null, error: { code: -32603, message: err.message } };
  }
}

const handleJsonRpc = async (req: Request, res: Response): Promise<void> => {
  const body = req.body;
  if (!body || (typeof body !== 'object')) {
    res.status(400).json({ jsonrpc: '2.0', error: { code: -32700, message: 'Parse error' } });
    return;
  }
  if (Array.isArray(body)) {
    const results = await Promise.all(body.map(processJsonRpcRequest));
    res.json(results);
    return;
  }
  const result = await processJsonRpcRequest(body);
  res.json(result);
};

app.post('/', jsonParser, handleJsonRpc);
app.post('/mcp', jsonParser, handleJsonRpc);
app.post('/api', jsonParser, handleJsonRpc);
app.post('/tools', jsonParser, handleJsonRpc);

// =============================================================================
// Start
// =============================================================================

console.log('[Init] Environment variable check:');
console.log(`[Init]   POWER_PLATFORM_ENVIRONMENT_ID = ${process.env.POWER_PLATFORM_ENVIRONMENT_ID ? '"' + process.env.POWER_PLATFORM_ENVIRONMENT_ID + '"' : '(not set)'}`);
console.log(`[Init]   AZURE_TENANT_ID = ${process.env.AZURE_TENANT_ID ? '(set)' : '(not set)'}`);
console.log(`[Init]   DATAVERSE_URL = ${process.env.DATAVERSE_URL ? '"' + process.env.DATAVERSE_URL + '"' : '(not set)'}`);

try {
  const defaultEnv = resolveEnvironmentId();
  console.log(`[Init] Default environment: ${defaultEnv}`);
} catch (e: any) {
  console.warn(`[Init] No default environment configured: ${e.message}`);
}

app.listen(PORT, () => {
  console.log('');
  console.log(`[Init] Power Automate MCP v${VERSION} running on port ${PORT}`);
  console.log(`[Init] SSE:       http://localhost:${PORT}/sse`);
  console.log(`[Init] REST:      http://localhost:${PORT}/mcp (+ /, /api, /tools)`);
  console.log(`[Init] Health:    http://localhost:${PORT}/health`);
  console.log(`[Init] API:       BAP admin + Flow admin + PowerApps admin (3 scopes)`);
  console.log(`[Init] Dataverse: ${solutionClient.isConfigured() ? process.env.DATAVERSE_URL + ' (configured)' : 'Not configured (set DATAVERSE_URL)'}`);
  console.log(`[Init] Write:     Delegated user token via Device Code Flow (multi-scope)`);
  console.log(`[Init] Auth:      ${userAuthManager.isConfigured() ? 'Dual-token mode (service principal + per-user delegated)' : 'Service-principal only (UserAuth not configured)'}`);
  console.log(`[Init] Tools:     ${toolDefs.length} registered (unified SSE + REST)`);
  console.log('');
});
