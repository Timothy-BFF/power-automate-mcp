// ═══════════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Main Entry Point
//
// Designed for Simtheory.ai integration via SSE transport.
// Deployed on Railway with health endpoint.
// Startup-resilient: boots even without Azure credentials.
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
// Only initializes if Azure AD credentials are fully configured.
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
      action: 'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in Railway environment variables, then redeploy.',
    });
  }
  return null;
}

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
  'Lists all Power Automate flows in a Power Platform environment. Returns flow name, display name, state (Started/Stopped), created time, and last modified time. Use filter to narrow by personal or shared flows.',
  listFlowsSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = listFlowsSchema.parse(args);
    const result = await executeListFlows(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 2. pa-get-flow-details
mcpServer.tool(
  'pa-get-flow-details',
  'Gets complete details about a specific Power Automate flow, including its definition, state, and HTTP trigger URI. Use pa-list-flows first to discover flow IDs.',
  getFlowDetailsSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = getFlowDetailsSchema.parse(args);
    const result = await executeGetFlowDetails(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 3. pa-enable-disable-flow
mcpServer.tool(
  'pa-enable-disable-flow',
  'Enables or disables a Power Automate flow. Use action "enable" to start a stopped flow, or "disable" to stop a running flow.',
  enableDisableFlowSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = enableDisableFlowSchema.parse(args);
    const result = await executeEnableDisableFlow(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 4. pa-delete-flow
mcpServer.tool(
  'pa-delete-flow',
  'Permanently deletes a Power Automate flow. DESTRUCTIVE AND IRREVERSIBLE. confirmDelete must be true. Verify with pa-get-flow-details first.',
  deleteFlowSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = deleteFlowSchema.parse(args);
    const result = await executeDeleteFlow(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 5. pa-trigger-flow
mcpServer.tool(
  'pa-trigger-flow',
  'Triggers a Power Automate flow with an HTTP Request trigger. Provide triggerUri directly or environmentId + flowId to auto-discover. Optionally pass a JSON body as input.',
  triggerFlowSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = triggerFlowSchema.parse(args);
    const result = await executeTriggerFlow(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 6. pa-get-run-history
mcpServer.tool(
  'pa-get-run-history',
  'Gets execution run history for a flow. Returns status, timing, trigger name, and error details. Use top to limit and status to filter.',
  getRunHistorySchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = getRunHistorySchema.parse(args);
    const result = await executeGetRunHistory(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 7. pa-get-run-details
mcpServer.tool(
  'pa-get-run-details',
  'Gets detailed information about a specific flow run including full status, timing, trigger info, and error details. Use pa-get-run-history to find run IDs.',
  getRunDetailsSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = getRunDetailsSchema.parse(args);
    const result = await executeGetRunDetails(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 8. pa-cancel-run
mcpServer.tool(
  'pa-cancel-run',
  'Cancels a currently running flow execution. Only works on runs in "Running" status. Use pa-get-run-history with status "Running" to find active runs.',
  cancelRunSchema.shape,
  async (args) => {
    const guard = requireConfigured();
    if (guard) return { content: [{ type: 'text' as const, text: guard }] };
    const parsed = cancelRunSchema.parse(args);
    const result = await executeCancelRun(parsed, flowClient!, defaultEnvId);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 9. pa-list-environments (no input parameters)
mcpServer.tool(
  'pa-list-environments',
  'Lists all Power Platform environments accessible to the service principal. Returns name, display name, location, type, state, and default flag. Use to discover environment IDs.',
  {},
  async () => {
    if (!config.azure.isConfigured || !envClient) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            error: 'Azure AD credentials are not configured.',
            action: 'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in Railway environment variables.',
          }),
        }],
      };
    }
    const result = await executeListEnvironments(envClient);
    return { content: [{ type: 'text' as const, text: result }] };
  }
);

// 10. pa-list-connections
mcpServer.tool(
  'pa-list-connections',
  'Lists all API connections in a Power Platform environment. Returns connector type, status, and creation time. Useful for auditing flow dependencies.',
  listConnectionsSchema.shape,
  async (args) => {
    if (!config.azure.isConfigured || !connClient) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            error: 'Azure AD credentials are not configured.',
            action: 'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in Railway environment variables.',
          }),
        }],
      };
    }
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

  // If no Simtheory token is configured, skip validation (development mode)
  if (!config.simtheoryAuthToken) {
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
  const missingVars: string[] = [];
  if (!config.azure.tenantId) missingVars.push('AZURE_TENANT_ID');
  if (!config.azure.clientId) missingVars.push('AZURE_CLIENT_ID');
  if (!config.azure.clientSecret) missingVars.push('AZURE_CLIENT_SECRET');
  if (!config.simtheoryAuthToken) missingVars.push('SIMTHEORY_AUTH_TOKEN');

  let tokenStatus = { flow: 'not_configured', management: 'not_configured' };

  if (tokenManager) {
    try {
      const health = await tokenManager.healthCheck();
      tokenStatus = {
        flow: health.flow ? 'ok' : 'error',
        management: health.management ? 'ok' : 'error',
      };
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
    tools: 10,
  });
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
  if (!config.azure.isConfigured) {
    logger.warn('Awaiting Azure AD configuration — add credentials to Railway variables and redeploy.');
  }
  logger.info('Waiting for Simtheory.ai connections...');
});
