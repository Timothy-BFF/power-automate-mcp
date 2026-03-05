// @ts-nocheck
// ═══════════════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Main Entry Point
//
// SIMTHEORY.AI PROTOCOL (fully mapped from production logs):
//
// Phase 1 — Discovery (REST JSON-RPC):
//   POST /        → initialize, tools/list
//   GET /         → server info
//   POST /tools   → tools/list
//   GET /tools    → REST tool definitions
//   POST /mcp     → tools/list
//   GET /mcp      → server info
//   POST /api     → tools/list
//   GET /api      → server info
//
// Phase 2 — SSE Transport:
//   GET /events   → SSE (Simtheory/mcp-operator)
//   GET /sse      → SSE (python-requests)
//
// Phase 3 — Tool Execution (REST JSON-RPC):
//   POST /sse     → tools/call
//   POST /mcp     → tools/call
//   POST /api     → tools/call
//   POST /        → tools/call
//
// Tool execution goes through direct REST JSON-RPC POST, NOT the
// SSE transport. The SSE channel is used for streaming/events only.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════════════════

import express from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { loadConfig, AppConfig } from './config/index.js';
import { createLogger } from './utils/logger.js';
import { AzureTokenManager } from './auth/azure-token-manager.js';
import { FlowClient } from './clients/flow-client.js';
import { EnvironmentClient } from './clients/environment-client.js';
import { ConnectionClient } from './clients/connection-client.js';

// Tool imports — schemas and executors
import { listFlowsSchema, executeListFlows } from './mcp/tools/list-flows.js';
import { getFlowDetailsSchema, executeGetFlowDetails } from './mcp/tools/get-flow-details.js';
import { enableDisableFlowSchema, executeEnableDisableFlow } from './mcp/tools/enable-disable-flow.js';
import { deleteFlowSchema, executeDeleteFlow } from './mcp/tools/delete-flow.js';
import { triggerFlowSchema, executeTriggerFlow } from './mcp/tools/trigger-flow.js';
import { getRunHistorySchema, executeGetRunHistory } from './mcp/tools/get-run-history.js';
import { getRunDetailsSchema, executeGetRunDetails } from './mcp/tools/get-run-details.js';
import { cancelRunSchema, executeCancelRun } from './mcp/tools/cancel-run.js';
import { executeListEnvironments } from './mcp/tools/list-environments.js';
import { listConnectionsSchema, executeListConnections } from './mcp/tools/list-connections.js';

// ─────────────────────────────────────────────────────────────────
// Bootstrap
// ─────────────────────────────────────────────────────────────────

const config: AppConfig = loadConfig();
const logger = createLogger(config.logLevel);

// ─────────────────────────────────────────────────────────────────
// Global Crash Protection
// ─────────────────────────────────────────────────────────────────

process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled Promise Rejection — server staying alive', {
    reason: reason instanceof Error ? reason.message : String(reason),
    stack: reason instanceof Error ? reason.stack : undefined,
  });
});

process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception — server staying alive', {
    message: error.message,
    stack: error.stack,
  });
});

logger.info('╔══════════════════════════════════════════════════════════╗');
logger.info('║   Power Automate MCP Server                             ║');
logger.info('║   GROW by Bolthouse Fresh (Architected by MCA)          ║');
logger.info('╚══════════════════════════════════════════════════════════╝');

// ─────────────────────────────────────────────────────────────────
// Initialize Azure Token Manager (No-Bother Protocol)
// ─────────────────────────────────────────────────────────────────

let tokenManager: AzureTokenManager | null = null;
let flowClient: FlowClient | null = null;
let envClient: EnvironmentClient | null = null;
let connClient: ConnectionClient | null = null;

if (config.azure.isConfigured) {
  logger.info('Azure AD credentials detected — initializing token manager.');
  tokenManager = AzureTokenManager.initialize(
    {
      tenantId: config.azure.tenantId,
      clientId: config.azure.clientId,
      clientSecret: config.azure.clientSecret,
      tokenEndpoint: config.azure.tokenEndpoint,
      scopes: {
        flow: config.azure.flowScope,
        management: config.azure.managementScope,
      },
    },
    logger
  );

  const flowHttpClient = tokenManager.createAuthenticatedClient('flow', config.powerPlatform.flowApiBase);
  const envHttpClient = tokenManager.createAuthenticatedClient('management', config.powerPlatform.environmentApiBase);
  const connHttpClient = tokenManager.createAuthenticatedClient('flow', config.powerPlatform.flowApiBase);

  flowClient = new FlowClient(flowHttpClient, logger);
  envClient = new EnvironmentClient(envHttpClient, logger);
  connClient = new ConnectionClient(connHttpClient, logger);
} else {
  logger.warn('═══════════════════════════════════════════════════════');
  logger.warn('  Azure AD credentials NOT configured.');
  logger.warn('  Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
  logger.warn('  in Railway environment variables, then redeploy.');
  logger.warn('  Server will start but tools will return errors.');
  logger.warn('═══════════════════════════════════════════════════════');
}

const defaultEnvId = config.powerPlatform.defaultEnvironmentId;

// ─────────────────────────────────────────────────────────────────
// Helper: guard for unconfigured state
// ─────────────────────────────────────────────────────────────────

function requireConfigured(): string | null {
  if (!config.azure.isConfigured || !flowClient) {
    return JSON.stringify({
      error: 'Azure AD credentials are not configured.',
      action: 'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in Railway environment variables and redeploy.',
    });
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════
// Tool Definitions (served via REST discovery)
// ═══════════════════════════════════════════════════════════════

const TOOL_DEFINITIONS = [
  {
    name: 'pa-list-environments',
    description: 'Lists all Power Platform environments accessible to the configured service principal. Returns environment ID, display name, location, SKU, and lifecycle state for each environment.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'pa-list-flows',
    description: 'Lists all Power Automate flows in a Power Platform environment. Returns flow name, display name, state (Started/Stopped), created time, and last modified time. Use filter to narrow by personal or shared flows. Provide environmentId or uses the default configured environment.',
    inputSchema: {
      type: 'object',
      properties: {
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
        filter: { type: 'string', enum: ['personal', 'shared', 'all'], description: 'Filter by ownership type: personal (my flows), shared (team flows), or all.' },
        top: { type: 'number', description: 'Maximum number of flows to return.' },
      },
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
  },
  {
    name: 'pa-enable-disable-flow',
    description: 'Enables or disables a Power Automate flow. Use action "start" to enable or "stop" to disable the flow.',
    inputSchema: {
      type: 'object',
      properties: {
        flowId: { type: 'string', description: 'The unique identifier of the flow.' },
        action: { type: 'string', enum: ['start', 'stop'], description: 'Action to perform: start (enable) or stop (disable) the flow.' },
        environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      },
      required: ['flowId', 'action'],
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
  },
];

const SERVER_INFO = { name: 'power-automate-mcp', version: '1.0.0' };

const MCP_CAPABILITIES = {
  protocolVersion: '2024-11-05',
  capabilities: { tools: { listChanged: true } },
  serverInfo: SERVER_INFO,
};

logger.info(`Tool definitions loaded: ${TOOL_DEFINITIONS.length} tools for REST discovery + execution.`);

// ═══════════════════════════════════════════════════════════════
// Tool Execution Router (REST JSON-RPC tools/call)
//
// Simtheory.ai calls tools via direct REST POST, not through the
// SSE transport. This router maps tool names to their executors.
// ═══════════════════════════════════════════════════════════════

async function executeToolCall(
  toolName: string,
  args: Record<string, any>
): Promise<{ content: Array<{ type: string; text: string }> }> {
  const configErr = requireConfigured();
  if (configErr) {
    return { content: [{ type: 'text', text: configErr }] };
  }

  let result: string;

  switch (toolName) {
    case 'pa-list-environments':
      result = await executeListEnvironments(envClient!);
      break;
    case 'pa-list-flows':
      result = await executeListFlows(args, flowClient!, defaultEnvId);
      break;
    case 'pa-get-flow-details':
      result = await executeGetFlowDetails(args, flowClient!, defaultEnvId);
      break;
    case 'pa-enable-disable-flow':
      result = await executeEnableDisableFlow(args, flowClient!, defaultEnvId);
      break;
    case 'pa-delete-flow':
      result = await executeDeleteFlow(args, flowClient!, defaultEnvId);
      break;
    case 'pa-trigger-flow':
      result = await executeTriggerFlow(args, flowClient!, defaultEnvId);
      break;
    case 'pa-get-run-history':
      result = await executeGetRunHistory(args, flowClient!, defaultEnvId);
      break;
    case 'pa-get-run-details':
      result = await executeGetRunDetails(args, flowClient!, defaultEnvId);
      break;
    case 'pa-cancel-run':
      result = await executeCancelRun(args, flowClient!, defaultEnvId);
      break;
    case 'pa-list-connections':
      result = await executeListConnections(args, connClient!, defaultEnvId);
      break;
    default:
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: `Unknown tool: ${toolName}` }) }],
      };
  }

  return { content: [{ type: 'text', text: result }] };
}

// ═══════════════════════════════════════════════════════════════
// Centralized JSON-RPC Handler
//
// Handles all MCP JSON-RPC methods:
//   - initialize       → server capabilities
//   - tools/list       → tool definitions
//   - tools/call       → tool execution (the critical one!)
//   - notifications/*  → acknowledgment
//
// Mounted on: /, /tools, /sse, /mcp, /api
// ═══════════════════════════════════════════════════════════════

async function handleJsonRpc(
  req: express.Request,
  res: express.Response
): Promise<boolean> {
  const body = req.body;

  // Only handle JSON-RPC 2.0 requests
  if (!body?.jsonrpc || body.jsonrpc !== '2.0' || !body.method) {
    return false; // Not a JSON-RPC request — let caller handle
  }

  const { method, id, params } = body;
  const endpoint = req.path;

  logger.info(`JSON-RPC: ${method} on ${endpoint}`, { id, toolName: params?.name });

  try {
    switch (method) {
      case 'initialize': {
        res.json({ jsonrpc: '2.0', id, result: MCP_CAPABILITIES });
        return true;
      }

      case 'tools/list': {
        logger.info(`Serving ${TOOL_DEFINITIONS.length} tool definitions via JSON-RPC`);
        res.json({ jsonrpc: '2.0', id, result: { tools: TOOL_DEFINITIONS } });
        return true;
      }

      case 'tools/call': {
        const toolName = params?.name;
        const toolArgs = params?.arguments || {};

        logger.info(`Tool call: ${toolName}`, { args: JSON.stringify(toolArgs).substring(0, 500) });

        try {
          const result = await executeToolCall(toolName, toolArgs);
          logger.info(`Tool call succeeded: ${toolName}`, {
            resultLength: result.content[0]?.text?.length || 0,
          });
          res.json({ jsonrpc: '2.0', id, result });
        } catch (execError) {
          const err = execError instanceof Error ? execError : new Error(String(execError));
          logger.error(`Tool call failed: ${toolName}`, { message: err.message, stack: err.stack });
          res.json({
            jsonrpc: '2.0',
            id,
            result: {
              content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }],
              isError: true,
            },
          });
        }
        return true;
      }

      // Handle all notification methods
      default: {
        if (method.startsWith('notifications/')) {
          res.json({ jsonrpc: '2.0', id: id || null, result: {} });
          return true;
        }

        logger.warn(`Unknown JSON-RPC method: ${method} on ${endpoint}`);
        res.json({
          jsonrpc: '2.0',
          id,
          error: { code: -32601, message: `Method not found: ${method}` },
        });
        return true;
      }
    }
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error(`JSON-RPC handler error for ${method}`, { message: err.message, stack: err.stack });
    res.json({
      jsonrpc: '2.0',
      id,
      error: { code: -32603, message: `Internal error: ${err.message}` },
    });
    return true;
  }
}

// ─────────────────────────────────────────────────────────────────
// MCP SDK Tool Registration (kept for SSE transport compatibility)
// ─────────────────────────────────────────────────────────────────

const mcpServer = new McpServer({ name: 'power-automate-mcp', version: '1.0.0' });
const registerTool = mcpServer.tool.bind(mcpServer);

registerTool('pa-list-environments', {}, async () => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeListEnvironments(envClient!);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-list-flows', listFlowsSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeListFlows(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-get-flow-details', getFlowDetailsSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeGetFlowDetails(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-enable-disable-flow', enableDisableFlowSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeEnableDisableFlow(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-delete-flow', deleteFlowSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeDeleteFlow(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-trigger-flow', triggerFlowSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeTriggerFlow(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-get-run-history', getRunHistorySchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeGetRunHistory(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-get-run-details', getRunDetailsSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeGetRunDetails(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-cancel-run', cancelRunSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeCancelRun(args, flowClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

registerTool('pa-list-connections', listConnectionsSchema.shape, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  const result = await executeListConnections(args, connClient!, defaultEnvId);
  return { content: [{ type: 'text', text: result }] };
});

logger.info('All 10 MCP tools registered (SDK + REST discovery + REST execution).');

// ═══════════════════════════════════════════════════════════════
// Express Server
// ═══════════════════════════════════════════════════════════════

const app = express();
app.use(express.json());

// ─────── Request Logging ───────
app.use((req, res, next) => {
  if (req.path !== '/health') {
    const logData: Record<string, any> = {
      hasAuth: !!req.headers['authorization'],
      userAgent: req.headers['user-agent'] || 'unknown',
      contentType: req.headers['content-type'] || 'none',
    };
    if (req.method === 'POST' && req.body) {
      logData.body = JSON.stringify(req.body).substring(0, 500);
    }
    logger.info(`Incoming: ${req.method} ${req.path}`, logData);
  }
  next();
});

// ═══════════════════════════════════════════════════════════════
// Endpoint: / (root)
// ═══════════════════════════════════════════════════════════════

app.get('/', (_req, res) => {
  logger.info('REST discovery: GET /');
  res.json({ jsonrpc: '2.0', result: MCP_CAPABILITIES });
});

app.post('/', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.json({ ...SERVER_INFO, tools: TOOL_DEFINITIONS.length });
  }
});

// ═══════════════════════════════════════════════════════════════
// Endpoint: /tools
// ═══════════════════════════════════════════════════════════════

app.get('/tools', (_req, res) => {
  logger.info('REST discovery: GET /tools');
  res.json({ tools: TOOL_DEFINITIONS });
});

app.post('/tools', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.json({ tools: TOOL_DEFINITIONS });
  }
});

// ═══════════════════════════════════════════════════════════════
// Endpoint: /mcp (Simtheory probes this)
// ═══════════════════════════════════════════════════════════════

app.get('/mcp', (_req, res) => {
  logger.info('REST discovery: GET /mcp');
  res.json({ jsonrpc: '2.0', result: MCP_CAPABILITIES });
});

app.post('/mcp', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.json({ ...SERVER_INFO, tools: TOOL_DEFINITIONS.length });
  }
});

// ═══════════════════════════════════════════════════════════════
// Endpoint: /api (Simtheory probes this)
// ═══════════════════════════════════════════════════════════════

app.get('/api', (_req, res) => {
  logger.info('REST discovery: GET /api');
  res.json({ jsonrpc: '2.0', result: MCP_CAPABILITIES });
});

app.post('/api', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.json({ ...SERVER_INFO, tools: TOOL_DEFINITIONS.length });
  }
});

// ═══════════════════════════════════════════════════════════════
// Health Endpoint
// ═══════════════════════════════════════════════════════════════

app.get('/health', async (_req, res) => {
  try {
    const missingVars: string[] = [];
    if (!config.azure.tenantId) missingVars.push('AZURE_TENANT_ID');
    if (!config.azure.clientId) missingVars.push('AZURE_CLIENT_ID');
    if (!config.azure.clientSecret) missingVars.push('AZURE_CLIENT_SECRET');
    if (!config.simtheoryAuthToken) missingVars.push('SIMTHEORY_AUTH_TOKEN');

    let tokenStatus = { flow: 'not_configured', management: 'not_configured' };
    if (tokenManager) {
      try {
        const health = await tokenManager.healthCheck();
        tokenStatus = { flow: health.flow ? 'ok' : 'error', management: health.management ? 'ok' : 'error' };
      } catch {
        tokenStatus = { flow: 'error', management: 'error' };
      }
    }

    res.json({
      status: config.azure.isConfigured ? 'healthy' : 'awaiting_configuration',
      server: 'power-automate-mcp',
      version: '1.0.0',
      timestamp: new Date().toISOString(),
      configuration: {
        azureAD: config.azure.isConfigured ? 'configured' : 'MISSING',
        simtheoryAuth: config.simtheoryAuthToken ? 'configured' : 'MISSING',
        defaultEnvironment: config.powerPlatform.defaultEnvironmentId || 'not_set',
        missingVariables: missingVars.length > 0 ? missingVars : undefined,
      },
      auth: tokenStatus,
      tools: TOOL_DEFINITIONS.length,
    });
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error('Health check failed', { message: err.message });
    res.status(503).json({ status: 'unhealthy', error: err.message });
  }
});

// ═══════════════════════════════════════════════════════════════
// SSE Transport Endpoints
// ═══════════════════════════════════════════════════════════════

const transports: Map<string, SSEServerTransport> = new Map();

async function handleSSEConnection(req: express.Request, res: express.Response, endpoint: string) {
  try {
    logger.info(`New SSE connection on ${endpoint}`);

    try {
      await mcpServer.close();
      logger.info('Previous MCP transport closed for reconnection');
    } catch {
      logger.debug('No previous transport to close');
    }

    for (const [id, oldTransport] of transports.entries()) {
      try { await oldTransport.close(); } catch { /* already closed */ }
      transports.delete(id);
    }

    const transport = new SSEServerTransport('/messages', res);
    transports.set(transport.sessionId, transport);
    logger.info(`SSE session created: ${transport.sessionId} (via ${endpoint})`);

    res.on('close', () => {
      logger.info(`SSE connection closed: ${transport.sessionId}`);
      transports.delete(transport.sessionId);
    });

    res.on('error', (err) => {
      logger.error(`SSE connection error: ${transport.sessionId}`, { message: err.message });
      transports.delete(transport.sessionId);
    });

    await mcpServer.connect(transport);
    logger.info(`MCP server connected to SSE session: ${transport.sessionId}`);
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error(`SSE connection setup failed on ${endpoint}`, { message: err.message, stack: err.stack });
    if (!res.headersSent) {
      res.status(500).json({ error: 'SSE connection failed' });
    }
  }
}

app.get('/sse', (req, res) => handleSSEConnection(req, res, '/sse'));
app.get('/events', (req, res) => handleSSEConnection(req, res, '/events'));

// POST /sse — Simtheory sends tools/call here via JSON-RPC
app.post('/sse', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.status(400).json({ error: 'Expected JSON-RPC 2.0 request' });
  }
});

// POST /events — Simtheory may also send tools/call here
app.post('/events', async (req, res) => {
  const handled = await handleJsonRpc(req, res);
  if (!handled) {
    res.status(400).json({ error: 'Expected JSON-RPC 2.0 request' });
  }
});

// SSE message handler (for MCP SDK transport)
app.post('/messages', async (req, res) => {
  try {
    const sessionId = req.query.sessionId as string;
    logger.info(`Message received for session: ${sessionId}`);

    const transport = transports.get(sessionId);
    if (!transport) {
      logger.warn(`Session not found: ${sessionId}. Active sessions: ${transports.size}`);
      res.status(404).json({ error: 'Session not found' });
      return;
    }

    await transport.handlePostMessage(req, res);
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error('Message handling failed', { message: err.message, stack: err.stack });
    if (!res.headersSent) {
      res.status(500).json({ error: 'Message handling failed' });
    }
  }
});

// ─────── Global Error Handler ───────
app.use((err: Error, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  logger.error('Express unhandled error', { message: err.message, stack: err.stack });
  if (!res.headersSent) {
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────── Start Server ───────
app.listen(config.port, '0.0.0.0', () => {
  logger.info(`Server listening on port ${config.port}`);
  logger.info(`Health:  http://0.0.0.0:${config.port}/health`);
  logger.info(`SSE:     http://0.0.0.0:${config.port}/sse`);
  logger.info(`Events:  http://0.0.0.0:${config.port}/events`);
  logger.info(`Tools:   http://0.0.0.0:${config.port}/tools`);
  logger.info(`MCP:     http://0.0.0.0:${config.port}/mcp`);
  logger.info(`API:     http://0.0.0.0:${config.port}/api`);
  logger.info('REST discovery + JSON-RPC execution endpoints active');
  if (!config.azure.isConfigured) {
    logger.warn('Awaiting Azure AD configuration.');
  }
  logger.info('Waiting for Simtheory.ai connections...');
});
