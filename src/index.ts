// ═══════════════════════════════════════════════════════════════════════════
// Power Automate MCP Server — Main Entry Point
//
// Designed for Simtheory.ai integration via SSE transport.
// Deployed on Railway with health endpoint.
// Startup-resilient: boots even without Azure credentials.
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
// Global Crash Protection — server must NEVER die on errors
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
      action: 'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in Railway environment variables and redeploy.',
    });
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────
// MCP Server + Tool Registration
// ─────────────────────────────────────────────────────────────────

const mcpServer = new McpServer({
  name: 'power-automate-mcp',
  version: '1.0.0',
});

mcpServer.tool('pa-list-environments', {}, async () => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeListEnvironments(envClient!);
});

mcpServer.tool('pa-list-flows', listFlowsSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeListFlows(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-get-flow-details', getFlowDetailsSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeGetFlowDetails(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-enable-disable-flow', enableDisableFlowSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeEnableDisableFlow(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-delete-flow', deleteFlowSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeDeleteFlow(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-trigger-flow', triggerFlowSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeTriggerFlow(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-get-run-history', getRunHistorySchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeGetRunHistory(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-get-run-details', getRunDetailsSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeGetRunDetails(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-cancel-run', cancelRunSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeCancelRun(flowClient!, args, defaultEnvId);
});

mcpServer.tool('pa-list-connections', listConnectionsSchema, async (args) => {
  const err = requireConfigured();
  if (err) return { content: [{ type: 'text', text: err }] };
  return executeListConnections(connClient!, args, defaultEnvId);
});

logger.info('All 10 MCP tools registered successfully.');

// ─────────────────────────────────────────────────────────────────
// Express Server + SSE Transport
// ─────────────────────────────────────────────────────────────────

const app = express();
app.use(express.json());

// ─────── Simtheory.ai Auth Middleware (with diagnostic logging) ───────
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
    logger.warn('Missing Authorization header from incoming request', {
      path: req.path,
      headersPresent: Object.keys(req.headers).join(', '),
    });
    res.status(401).json({ error: 'Missing Authorization header' });
    return;
  }

  // Extract token — supports both "Bearer <token>" and raw "<token>"
  // .trim() defends against invisible whitespace from either side
  const token = (authHeader.startsWith('Bearer ')
    ? authHeader.slice(7)
    : authHeader).trim();

  if (token !== config.simtheoryAuthToken) {
    // ── Diagnostic logging: show WHY the comparison failed ──
    logger.warn('Invalid Simtheory.ai authorization token received', {
      receivedLength: token.length,
      expectedLength: config.simtheoryAuthToken.length,
      receivedFirst4: token.substring(0, 4) + '...',
      expectedFirst4: config.simtheoryAuthToken.substring(0, 4) + '...',
      receivedLast4: '...' + token.substring(Math.max(0, token.length - 4)),
      expectedLast4: '...' + config.simtheoryAuthToken.substring(Math.max(0, config.simtheoryAuthToken.length - 4)),
      headerFormat: authHeader.startsWith('Bearer ') ? 'Bearer <token>' : 'raw',
      match: token === config.simtheoryAuthToken,
    });
    res.status(403).json({ error: 'Invalid authorization token' });
    return;
  }

  logger.info('Simtheory.ai token validated successfully');
  next();
}

app.use(validateSimtheoryToken);

// ─────── Health Endpoint ───────
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
  } catch (error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error('Health check failed', { message: err.message });
    res.status(503).json({ status: 'unhealthy', error: err.message });
  }
});

// ─────── SSE Transport for MCP ───────
const transports: Map<string, SSEServerTransport> = new Map();

app.get('/sse', async (req, res) => {
  try {
    logger.info('New SSE connection request received');
    const transport = new SSEServerTransport('/messages', res);
    transports.set(transport.sessionId, transport);
    logger.info(`SSE session created: ${transport.sessionId}`);

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
  } catch (error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error('SSE connection setup failed', { message: err.message, stack: err.stack });
    if (!res.headersSent) {
      res.status(500).json({ error: 'SSE connection failed' });
    }
  }
});

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
  } catch (error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    logger.error('Message handling failed', { message: err.message, stack: err.stack });
    if (!res.headersSent) {
      res.status(500).json({ error: 'Message handling failed' });
    }
  }
});

// ─────── Global Express Error Handler ───────
app.use((err: Error, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  logger.error('Express unhandled error', { message: err.message, stack: err.stack });
  if (!res.headersSent) {
    res.status(500).json({ error: 'Internal server error' });
  }
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
