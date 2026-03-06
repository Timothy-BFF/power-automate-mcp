import 'dotenv/config';
import express, { Request, Response } from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { z } from 'zod';
import { AzureTokenManager } from './auth/azure-token-manager.js';
import { PowerPlatformClient } from './api/power-platform-client.js';
import { resolveEnvironmentId } from './config/environment-resolver.js';
import { ToolResult, ToolDefinition } from './types.js';

// =============================================================================
// Configuration
// =============================================================================
const PORT = parseInt(process.env.PORT || '8080', 10);
const VERSION = '2.2.1';

// =============================================================================
// Core Services
// =============================================================================
const tokenManager = new AzureTokenManager();
const client = new PowerPlatformClient(tokenManager);

// Pre-warm tokens (BAP + Flow + PowerApps scopes)
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
// =============================================================================
const toolDefs: ToolDefinition[] = [
  {
    name: 'pa-list-environments',
    description: 'Lists all Power Platform environments accessible to the configured service principal. Returns environment ID, display name, location, SKU, and lifecycle state for each environment.',
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
  {
    name: 'pa-list-flows',
    description: 'Lists all Power Automate flows in a Power Platform environment. Returns flow name, display name, state (Started/Stopped), created time, and last modified time. Use filter to narrow by personal or shared flows. Provide environmentId or uses the default configured environment.',
    inputSchema: {
      type: 'object',
      properties: {
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
        filter: { type: 'string', enum: ['personal', 'shared', 'all'], description: 'Filter by ownership type.' },
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
        return ok({ count: flows.length, environmentId: envId, flows });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-get-flow-details',
    description: 'Gets detailed information about a specific Power Automate flow including its definition, triggers, actions, connections, and current state.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId) return fail('flowId is required');
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.getFlowDetails(envId, p.flowId);
        return ok({
          id: r.name, displayName: r.properties?.displayName, state: r.properties?.state,
          createdTime: r.properties?.createdTime, lastModifiedTime: r.properties?.lastModifiedTime,
          definition: r.properties?.definition, connectionReferences: r.properties?.connectionReferences,
          triggers: r.properties?.definition?.triggers, actions: r.properties?.definition?.actions,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-enable-disable-flow',
    description: 'Enables or disables a Power Automate flow. Use action "start" to enable or "stop" to disable the flow.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        action: { type: 'string', enum: ['start', 'stop'], description: 'Action to perform.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId', 'action'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId || !p.action) return fail('flowId and action are required');
        const envId = resolveEnvironmentId(p.environmentId);
        await client.enableDisableFlow(envId, p.flowId, p.action);
        return ok({ status: 'success', flowId: p.flowId, action: p.action });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-delete-flow',
    description: 'Permanently deletes a Power Automate flow. This action cannot be undone. Use with caution.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow to delete.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId) return fail('flowId is required');
        const envId = resolveEnvironmentId(p.environmentId);
        await client.deleteFlow(envId, p.flowId);
        return ok({ status: 'deleted', flowId: p.flowId });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-trigger-flow',
    description: 'Manually triggers a Power Automate flow that has an HTTP request trigger. Optionally pass a JSON body to the trigger.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow to trigger.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
        triggerBody: { type: 'object', description: 'Optional JSON body to pass to the flow trigger.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId) return fail('flowId is required');
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.triggerFlow(envId, p.flowId, p.triggerBody);
        return ok({ status: 'triggered', flowId: p.flowId, result: r });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-get-run-history',
    description: 'Gets the run history for a specific Power Automate flow. Returns run ID, status (Succeeded/Failed/Running/Cancelled), start time, end time, and trigger information.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
        top: { type: 'number', description: 'Maximum number of runs to return.' },
      },
      required: ['flowId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId) return fail('flowId is required');
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.getRunHistory(envId, p.flowId, p.top ? Number(p.top) : undefined);
        const runs = (r.value || []).map((run: any) => ({
          id: run.name, status: run.properties?.status,
          startTime: run.properties?.startTime, endTime: run.properties?.endTime,
          trigger: run.properties?.trigger?.name,
        }));
        return ok({ count: runs.length, flowId: p.flowId, runs });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-get-run-details',
    description: 'Gets detailed information about a specific flow run including action-level results, inputs, outputs, and timing.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        runId: { type: 'string', description: 'The unique identifier of the run.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId', 'runId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId || !p.runId) return fail('flowId and runId are required');
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.getRunDetails(envId, p.flowId, p.runId);
        return ok({
          id: r.name, status: r.properties?.status,
          startTime: r.properties?.startTime, endTime: r.properties?.endTime,
          trigger: r.properties?.trigger, outputs: r.properties?.outputs, error: r.properties?.error,
        });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-cancel-run',
    description: 'Cancels a currently running Power Automate flow run.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        runId: { type: 'string', description: 'The unique identifier of the run to cancel.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId', 'runId'],
    },
    handler: async (p: any) => {
      try {
        if (!p.flowId || !p.runId) return fail('flowId and runId are required');
        const envId = resolveEnvironmentId(p.environmentId);
        await client.cancelRun(envId, p.flowId, p.runId);
        return ok({ status: 'cancelled', flowId: p.flowId, runId: p.runId });
      } catch (e: any) { return fail(e.message); }
    },
  },
  {
    name: 'pa-list-connections',
    description: 'Lists all Power Platform connections in an environment. Returns connection ID, display name, status, connector information, and creation time.',
    inputSchema: {
      type: 'object',
      properties: {
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
    },
    handler: async (p: any) => {
      try {
        const envId = resolveEnvironmentId(p.environmentId);
        const r = await client.listConnections(envId);
        const conns = (r.value || []).map((c: any) => ({
          id: c.name, displayName: c.properties?.displayName,
          status: c.properties?.statuses, apiId: c.properties?.apiId,
          createdTime: c.properties?.createdTime,
        }));
        return ok({ count: conns.length, environmentId: envId, connections: conns });
      } catch (e: any) { return fail(e.message); }
    },
  },
];

// Build handler lookup map for REST JSON-RPC
const toolHandlers = new Map<string, (params: any) => Promise<ToolResult>>();
for (const t of toolDefs) {
  toolHandlers.set(t.name, t.handler);
}

// =============================================================================
// MCP Server (SSE transport for discovery + streaming)
// =============================================================================
const mcpServer = new McpServer({ name: 'power-automate-mcp', version: VERSION });

// Relaxed-type binding: bypasses strict TS overload resolution for MCP SDK v1.12+
const registerTool: any = mcpServer.tool.bind(mcpServer);

registerTool('pa-list-environments', toolDefs[0].description, {},
  async () => toolDefs[0].handler({}));

registerTool('pa-list-flows', toolDefs[1].description, {
  environmentId: z.string().optional().describe('Power Platform environment ID'),
  filter: z.enum(['personal', 'shared', 'all']).optional().describe('Filter by ownership'),
  top: z.number().optional().describe('Max flows to return'),
}, async (p: any) => toolDefs[1].handler(p));

registerTool('pa-get-flow-details', toolDefs[2].description, {
  flowId: z.string().describe('Flow ID'),
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[2].handler(p));

registerTool('pa-enable-disable-flow', toolDefs[3].description, {
  flowId: z.string().describe('Flow ID'),
  action: z.enum(['start', 'stop']).describe('Enable or disable'),
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[3].handler(p));

registerTool('pa-delete-flow', toolDefs[4].description, {
  flowId: z.string().describe('Flow ID'),
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[4].handler(p));

registerTool('pa-trigger-flow', toolDefs[5].description, {
  flowId: z.string().describe('Flow ID'),
  environmentId: z.string().optional().describe('Environment ID'),
  triggerBody: z.record(z.any()).optional().describe('Trigger body'),
}, async (p: any) => toolDefs[5].handler(p));

registerTool('pa-get-run-history', toolDefs[6].description, {
  flowId: z.string().describe('Flow ID'),
  environmentId: z.string().optional().describe('Environment ID'),
  top: z.number().optional().describe('Max runs'),
}, async (p: any) => toolDefs[6].handler(p));

registerTool('pa-get-run-details', toolDefs[7].description, {
  flowId: z.string().describe('Flow ID'),
  runId: z.string().describe('Run ID'),
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[7].handler(p));

registerTool('pa-cancel-run', toolDefs[8].description, {
  flowId: z.string().describe('Flow ID'),
  runId: z.string().describe('Run ID'),
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[8].handler(p));

registerTool('pa-list-connections', toolDefs[9].description, {
  environmentId: z.string().optional().describe('Environment ID'),
}, async (p: any) => toolDefs[9].handler(p));

console.log(`[Init] MCP tools registered: ${toolDefs.length}`);

// =============================================================================
// Express Server
// =============================================================================
const app = express();

// CORS middleware
app.use((req: Request, res: Response, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Mcp-Session-Id, Accept');
  if (req.method === 'OPTIONS') { res.sendStatus(200); return; }
  next();
});

const jsonParser = express.json();

// --- Health ---
app.get('/health', (_req: Request, res: Response) => {
  res.json({
    status: 'healthy', server: 'power-automate-mcp', version: VERSION,
    tools: toolDefs.length, uptime: Math.floor(process.uptime()),
    auth: {
      azureAD: process.env.AZURE_CLIENT_ID ? 'configured' : 'missing',
      simtheoryAuth: process.env.SIMTHEORY_AUTH_TOKEN ? 'configured' : 'missing',
    },
  });
});

// --- GET Discovery ---
app.get('/', (_req: Request, res: Response) => {
  res.json({
    name: 'power-automate-mcp', version: VERSION, protocol: 'MCP',
    transport: ['sse', 'json-rpc'],
    endpoints: { sse: '/sse', messages: '/messages', health: '/health', jsonRpc: ['/', '/mcp', '/api'] },
  });
});

app.get('/tools', (_req: Request, res: Response) => {
  res.json({ tools: toolDefs.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })) });
});

app.get('/mcp', (_req: Request, res: Response) => {
  res.json({ name: 'power-automate-mcp', version: VERSION, status: 'ready' });
});

app.get('/api', (_req: Request, res: Response) => {
  res.json({ name: 'power-automate-mcp', version: VERSION, status: 'ready' });
});

// --- SSE Transport ---
const sseTransports = new Map<string, SSEServerTransport>();

app.get('/sse', async (req: Request, res: Response) => {
  console.log('[SSE] Connection initiated');
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

// Diagnostic: Log environment variable resolution
console.log('[Init] Environment variable check:');
console.log(`[Init]   POWER_PLATFORM_ENVIRONMENT_ID = ${process.env.POWER_PLATFORM_ENVIRONMENT_ID ? '"' + process.env.POWER_PLATFORM_ENVIRONMENT_ID + '"' : '(not set)'}`);
console.log(`[Init]   DEFAULT_ENVIRONMENT_ID = ${process.env.DEFAULT_ENVIRONMENT_ID ? '"' + process.env.DEFAULT_ENVIRONMENT_ID + '"' : '(not set)'}`);
console.log(`[Init]   AZURE_TENANT_ID = ${process.env.AZURE_TENANT_ID ? '(set)' : '(not set)'}`);

try {
  const defaultEnv = resolveEnvironmentId();
  console.log(`[Init] Default environment: ${defaultEnv}`);
} catch (e: any) {
  console.warn(`[Init] No default environment configured: ${e.message}`);
}

app.listen(PORT, () => {
  console.log('');
  console.log(`[Init] Power Automate MCP v${VERSION} running on port ${PORT}`);
  console.log(`[Init] SSE:    http://localhost:${PORT}/sse`);
  console.log(`[Init] REST:   http://localhost:${PORT}/mcp (+ /, /api, /tools)`);
  console.log(`[Init] Health: http://localhost:${PORT}/health`);
  console.log(`[Init] API:    BAP admin + Flow admin (/scopes/admin/) + PowerApps admin`);
  console.log('');
});
