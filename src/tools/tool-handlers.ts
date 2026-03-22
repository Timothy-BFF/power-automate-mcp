/**
 * Power Automate MCP - Tool Handlers (v3.0.3)
 *
 * Registers all non-auth MCP tools with the McpServer.
 * Descriptions are imported from tool-descriptions.ts to provide
 * procedure guidance that prevents destructive agent behavior.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { PowerPlatformClient } from '../api/power-platform-client.js';
import { TOOL_DESCRIPTIONS } from './tool-descriptions.js';

// =========================================================================
// Response Helpers
// =========================================================================

function jsonResponse(data: any): { content: Array<{ type: 'text'; text: string }> } {
  // Strip _raw from getFlowDetails to keep response manageable
  if (data && data._raw) {
    const { _raw, ...clean } = data;
    return { content: [{ type: 'text' as const, text: JSON.stringify(clean, null, 2) }] };
  }
  return { content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }] };
}

function errorResponse(error: any): { content: Array<{ type: 'text'; text: string }>; isError: true } {
  const msg = error?.message || String(error);
  console.error(`[Tool] Error: ${msg}`);
  return { content: [{ type: 'text' as const, text: `Error: ${msg}` }], isError: true };
}

// =========================================================================
// Environment Resolver
// =========================================================================

function createEnvResolver(defaultEnvId: string) {
  return function resolveEnv(envId?: string): string {
    if (envId) {
      console.log(`[EnvResolver] Using explicit parameter: ${envId}`);
      return envId;
    }
    console.log(`[EnvResolver] Using POWER_PLATFORM_ENVIRONMENT_ID: ${defaultEnvId}`);
    return defaultEnvId;
  };
}

// =========================================================================
// Tool Registration
// =========================================================================

export function registerTools(
  server: McpServer,
  client: PowerPlatformClient,
  defaultEnvId: string
): number {
  const resolveEnv = createEnvResolver(defaultEnvId);
  let count = 0;

  // -----------------------------------------------------------------------
  // ENVIRONMENT TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-list-environments',
    TOOL_DESCRIPTIONS['pa-list-environments'],
    {},
    async () => {
      try {
        const result = await client.listEnvironments();
        const envs = result?.value || [];
        const summary = envs.map((e: any) => ({
          id: e.name,
          displayName: e.properties?.displayName,
          type: e.properties?.environmentType,
          state: e.properties?.states?.management?.id,
          region: e.location,
          createdTime: e.properties?.createdTime,
        }));
        return jsonResponse({ environments: summary, count: summary.length });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // FLOW READ TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-list-flows',
    TOOL_DESCRIPTIONS['pa-list-flows'],
    {
      environmentId: z.string().optional().describe('Power Platform environment ID (defaults to configured env)'),
      filter: z.string().optional().describe('OData filter expression'),
      top: z.number().optional().describe('Max number of flows to return'),
    },
    async ({ environmentId, filter, top }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.listFlows(envId, filter, top);
        const flows = (result?.value || []).map((f: any) => ({
          flowId: f.name,
          displayName: f.properties?.displayName,
          state: f.properties?.state,
          createdTime: f.properties?.createdTime,
          lastModifiedTime: f.properties?.lastModifiedTime,
          creator: f.properties?.creator?.userId || f.properties?.creator?.objectId,
        }));
        return jsonResponse({ flows, count: flows.length });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-get-flow-details',
    TOOL_DESCRIPTIONS['pa-get-flow-details'],
    {
      flowId: z.string().describe('Flow ID to retrieve details for'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.getFlowDetails(envId, flowId);
        return jsonResponse(result);
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // FLOW WRITE TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-create-flow',
    TOOL_DESCRIPTIONS['pa-create-flow'],
    {
      displayName: z.string().describe('Flow display name'),
      definition: z.any().describe('Complete workflow definition JSON with triggers and actions'),
      state: z.string().optional().describe('Initial state: "Started" or "Stopped" (default: Stopped)'),
      connectionReferences: z.any().optional().describe('Connection references for connectors'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ displayName, definition, state, connectionReferences, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.createFlow(envId, displayName, definition, state, connectionReferences);
        return jsonResponse(result);
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-update-flow',
    TOOL_DESCRIPTIONS['pa-update-flow'],
    {
      flowId: z.string().describe('Flow ID to update'),
      displayName: z.string().optional().describe('New display name'),
      definition: z.any().optional().describe('Updated workflow definition JSON'),
      state: z.string().optional().describe('New state: "Started" or "Stopped"'),
      connectionReferences: z.any().optional().describe('Updated connection references'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, displayName, definition, state, connectionReferences, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const updates: any = {};
        if (displayName) updates.displayName = displayName;
        if (definition) updates.definition = definition;
        if (state) updates.state = state;
        if (connectionReferences) updates.connectionReferences = connectionReferences;
        const result = await client.updateFlow(envId, flowId, updates);
        return jsonResponse(result);
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // FLOW LIFECYCLE TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-enable-flow',
    TOOL_DESCRIPTIONS['pa-enable-flow'],
    {
      flowId: z.string().describe('Flow ID to enable'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        await client.enableDisableFlow(envId, flowId, 'start');
        return jsonResponse({ status: 'enabled', flowId, message: `Flow ${flowId} has been enabled (started).` });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-disable-flow',
    TOOL_DESCRIPTIONS['pa-disable-flow'],
    {
      flowId: z.string().describe('Flow ID to disable'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        await client.enableDisableFlow(envId, flowId, 'stop');
        return jsonResponse({ status: 'disabled', flowId, message: `Flow ${flowId} has been disabled (stopped).` });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-delete-flow',
    TOOL_DESCRIPTIONS['pa-delete-flow'],
    {
      flowId: z.string().describe('Flow ID to delete'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        await client.deleteFlow(envId, flowId);
        return jsonResponse({ status: 'deleted', flowId, message: `Flow ${flowId} has been permanently deleted.` });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-trigger-flow',
    TOOL_DESCRIPTIONS['pa-trigger-flow'],
    {
      flowId: z.string().describe('Flow ID to trigger'),
      body: z.any().optional().describe('Request body for the trigger (for Request-triggered flows)'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, body, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.triggerFlow(envId, flowId, body);
        return jsonResponse({ status: 'triggered', flowId, result });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // FLOW RUN TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-get-run-history',
    TOOL_DESCRIPTIONS['pa-get-run-history'],
    {
      flowId: z.string().describe('Flow ID to get run history for'),
      top: z.number().optional().describe('Max number of runs to return'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, top, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.getRunHistory(envId, flowId, top);
        const runs = (result?.value || []).map((r: any) => ({
          runId: r.name,
          status: r.properties?.status,
          startTime: r.properties?.startTime,
          endTime: r.properties?.endTime,
          trigger: r.properties?.trigger?.name,
        }));
        return jsonResponse({ runs, count: runs.length, flowId });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-get-run-details',
    TOOL_DESCRIPTIONS['pa-get-run-details'],
    {
      flowId: z.string().describe('Flow ID'),
      runId: z.string().describe('Run ID to get details for'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, runId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.getRunDetails(envId, flowId, runId);
        return jsonResponse(result);
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  server.tool(
    'pa-cancel-run',
    TOOL_DESCRIPTIONS['pa-cancel-run'],
    {
      flowId: z.string().describe('Flow ID'),
      runId: z.string().describe('Run ID to cancel'),
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ flowId, runId, environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        await client.cancelRun(envId, flowId, runId);
        return jsonResponse({ status: 'cancelled', flowId, runId, message: `Run ${runId} has been cancelled.` });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // CONNECTION TOOLS
  // -----------------------------------------------------------------------

  server.tool(
    'pa-list-connections',
    TOOL_DESCRIPTIONS['pa-list-connections'],
    {
      environmentId: z.string().optional().describe('Power Platform environment ID'),
    },
    async ({ environmentId }) => {
      try {
        const envId = resolveEnv(environmentId);
        const result = await client.listConnections(envId);
        const connections = (result?.value || []).map((c: any) => ({
          connectionId: c.name,
          displayName: c.properties?.displayName,
          apiId: c.properties?.apiId,
          status: c.properties?.statuses?.[0]?.status,
          createdTime: c.properties?.createdTime,
        }));
        return jsonResponse({ connections, count: connections.length });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  console.log(`[Init] MCP tools registered: ${count}`);
  return count;
}
