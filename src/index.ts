// ═══════════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Main Entry Point
// 
// Designed for Simtheory.ai integration via SSE transport.
// Deployed on Railway with health endpoint.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════════════

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

logger.info('╔══════════════════════════════════════════════════════════╗');
logger.info('║   Power Automate MCP Server                             ║');
logger.info('║   GROW by Bolthouse Fresh (Architected by MCA)          ║');
logger.info('╚══════════════════════════════════════════════════════════╝');

// ─────────────────────────────────────────────────────────────────
// Initialize Azure Token Manager (No-Bother Protocol)
// ─────────────────────────────────────────────────────────────────

const tokenManager = AzureTokenManager.initialize(
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

// ─────────────────────────────────────────────────────────────────
// Create Authenticated API Clients
// ─────────────────────────────────────────────────────────────────

const flowHttpClient = tokenManager.createAuthenticatedClient(
  'flow',
  config.powerPlatform.flowApiBase
);
const envHttpClient = tokenManager.createAuthenticatedClient(
  'management',
  config.powerPlatform.environmentApiBase
);
const connHttpClient = tokenManager.createAuthenticatedClient(
  'flow',
  config.powerPlatform.flowApiBase
);

const flowClient = new FlowClient(flowHttpClient, logger);
const envClient = new EnvironmentClient(envHttpClient, logger);
const connClient = new ConnectionClient(connHttpClient, logger);

const defaultEnvId = config.powerPlatform.defaultEnvironmentId;

// ─────────────────────────────────────────────────────────────────
// MCP Server Setup
// ─────────────────────────────────────────────────────────────────

const mcpServer = new McpServer({
  name: 'power-automate-mcp',
  version: '1.0.0',
});

// ─────────────────────────────────────────────────────────────────
// Register Tools (Zod shapes for MCP SDK v1.12+)
// ─────────────────────────────────────────────────────────────────

// 1. pa-list-flows
mcpServer.tool(
  'pa-list-flows',
  'Lists all Power Automate flows in a Power Platform environment. Returns flow name, display name, state (Started/Stopped), created time, and last modified time. Use filter to narrow by personal or shared flows. Provide environmentId or uses the default configured environment.',
  listFlowsSchema.shape,
  async (args) => {
    const parsed = listFlowsSchema.parse(args);
    const result = await executeListFlows(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 2. pa-get-flow-details
mcpServer.tool(
  'pa-get-flow-details',
  'Gets complete details about a specific Power Automate flow, including its definition (triggers, actions, conditions), state, creation time, and HTTP trigger URI if applicable. Use pa-list-flows first to discover flow IDs.',
  getFlowDetailsSchema.shape,
  async (args) => {
    const parsed = getFlowDetailsSchema.parse(args);
    const result = await executeGetFlowDetails(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 3. pa-enable-disable-flow
mcpServer.tool(
  'pa-enable-disable-flow',
  'Enables or disables a Power Automate flow. Use action "enable" to start a stopped flow, or "disable" to stop a running flow. Returns the new state after the operation.',
  enableDisableFlowSchema.shape,
  async (args) => {
    const parsed = enableDisableFlowSchema.parse(args);
    const result = await executeEnableDisableFlow(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 4. pa-delete-flow
mcpServer.tool(
  'pa-delete-flow',
  'Permanently deletes a Power Automate flow. THIS IS DESTRUCTIVE AND IRREVERSIBLE. The confirmDelete parameter must be set to true to proceed. Always verify the flow ID with pa-get-flow-details before deleting.',
  deleteFlowSchema.shape,
  async (args) => {
    const parsed = deleteFlowSchema.parse(args);
    const result = await executeDeleteFlow(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 5. pa-trigger-flow
mcpServer.tool(
  'pa-trigger-flow',
  'Triggers a Power Automate flow that has an HTTP Request trigger. Provide either the triggerUri directly (from pa-get-flow-details), or provide environmentId + flowId to auto-discover the trigger URI. Optionally pass a JSON body as input parameters to the flow. Returns the HTTP status code and response body.',
  triggerFlowSchema.shape,
  async (args) => {
    const parsed = triggerFlowSchema.parse(args);
    const result = await executeTriggerFlow(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 6. pa-get-run-history
mcpServer.tool(
  'pa-get-run-history',
  'Gets the execution run history for a specific Power Automate flow. Returns run status (Succeeded, Failed, Running, Cancelled), start time, end time, trigger name, and error details for failed runs. Use top to limit results and status to filter by outcome.',
  getRunHistorySchema.shape,
  async (args) => {
    const parsed = getRunHistorySchema.parse(args);
    const result = await executeGetRunHistory(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 7. pa-get-run-details
mcpServer.tool(
  'pa-get-run-details',
  'Gets detailed information about a specific Power Automate flow run, including full status, timing, trigger info, and error details. Use pa-get-run-history first to discover run IDs.',
  getRunDetailsSchema.shape,
  async (args) => {
    const parsed = getRunDetailsSchema.parse(args);
    const result = await executeGetRunDetails(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 8. pa-cancel-run
mcpServer.tool(
  'pa-cancel-run',
  'Cancels a currently running Power Automate flow execution. Only works on runs that are in "Running" status. Use pa-get-run-history with status filter "Running" to find active runs.',
  cancelRunSchema.shape,
  async (args) => {
    const parsed = cancelRunSchema.parse(args);
    const result = await executeCancelRun(parsed, flowClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 9. pa-list-environments (no input parameters)
mcpServer.tool(
  'pa-list-environments',
  'Lists all Power Platform environments accessible to the connected service principal. Returns environment name, display name, location, type (Production/Sandbox/Developer), state, and whether it is the default environment. Use this to discover environment IDs for other tools.',
  {},
  async () => {
    const result = await executeListEnvironments(envClient);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 10. pa-list-connections
mcpServer.tool(
  'pa-list-connections',
  'Lists all API connections configured in a Power Platform environment. Returns connection name, connector type (e.g., Office 365, SharePoint, SQL), status, and creation time. Useful for auditing or debugging flow dependencies.',
  listConnectionsSchema.shape,
  async (args) => {
    const parsed = listConnectionsSchema.parse(args);
    const result = await executeListConnections(parsed, connClient, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

logger.info('All 10 MCP tools registered successfully.');

// ─────────────────────────────────────────────────────────────────
// Express Server + SSE Transport
// ─────────────────────────────────────────────────────────────────

const app = express();
app.use(express.json());

// ─────── Simtheory.ai Auth Middleware ───────
function validateSimtheoryToken(
  req: express.Request,
  res: express.Response,
  next: express.NextFunction
): void {
  // Skip auth for health check
  if (req.path === '/health') {
    next();
    return;
  }

  const authHeader = req.headers['authorization'];
  if (!authHeader) {
    res.status(401).json({ error: 'Missing Authorization header' });
    return;
  }

  const token = authHeader.startsWith('Bearer ')
    ? authHeader.slice(7)
    : authHeader;

  if (token !== config.simtheoryAuthToken) {
    logger.warn('Invalid Simtheory.ai authorization token received');
    res.status(403).json({ error: 'Invalid authorization token' });
    return;
  }

  next();
}

app.use(validateSimtheoryToken);

// ─────── Health Endpoint ───────
app.get('/health', async (_req, res) => {
  try {
    const tokenHealth = await tokenManager.healthCheck();
    res.json({
      status: 'healthy',
      server: 'power-automate-mcp',
      version: '1.0.0',
      timestamp: new Date().toISOString(),
      auth: {
        flowToken: tokenHealth.flow ? 'ok' : 'unavailable',
        managementToken: tokenHealth.management ? 'ok' : 'unavailable',
      },
      tools: 10,
    });
  } catch {
    res.status(503).json({ status: 'unhealthy' });
  }
});

// ─────── SSE Transport for MCP ───────
const transports: Map<string, SSEServerTransport> = new Map();

app.get('/sse', async (req, res) => {
  logger.info('New SSE connection established');
  const transport = new SSEServerTransport('/messages', res);
  transports.set(transport.sessionId, transport);

  res.on('close', () => {
    logger.info(`SSE connection closed: ${transport.sessionId}`);
    transports.delete(transport.sessionId);
  });

  await mcpServer.connect(transport);
});

app.post('/messages', async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const transport = transports.get(sessionId);

  if (!transport) {
    res.status(404).json({ error: 'Session not found' });
    return;
  }

  await transport.handlePostMessage(req, res);
});

// ─────── Start Server ───────
app.listen(config.port, '0.0.0.0', () => {
  logger.info(`Server listening on port ${config.port}`);
  logger.info(`Health: http://0.0.0.0:${config.port}/health`);
  logger.info(`SSE:    http://0.0.0.0:${config.port}/sse`);
  logger.info('Waiting for Simtheory.ai connections...');
});
