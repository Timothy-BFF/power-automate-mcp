import express from 'express';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { createLogger, format, transports } from 'winston';
import { config } from './config/index.js';
import { TokenManager } from './auth/azure-token-manager.js';
import { PowerPlatformClient } from './api/power-platform-client.js';
import { createAndConfigureMcpServer } from './mcp/server.js';

const logger = createLogger({
  level: config.logLevel,
  format: format.combine(format.timestamp(), format.json()),
  transports: [new transports.Console()],
});

if (!config.tenantId || !config.clientId || !config.clientSecret) {
  logger.warn('Azure credentials not fully configured - API calls will fail');
}

const tokenManager = new TokenManager(config.tenantId, config.clientId, config.clientSecret, logger);
const apiClient = new PowerPlatformClient(tokenManager, logger);
const mcpServer = createAndConfigureMcpServer(apiClient, config.defaultEnvironmentId, logger);

const app = express();
app.use(express.json());

app.get('/health', (_req: any, res: any) => {
  res.json({ status: 'ok', server: 'power-automate-mcp', version: '2.0.0', timestamp: new Date().toISOString() });
});

const checkAuth = (req: any, res: any, next: any): void => {
  if (config.simtheoryToken) {
    const token = req.headers?.['authorization']?.replace('Bearer ', '').trim();
    if (token && token !== config.simtheoryToken) {
      res.status(401).json({ error: 'Unauthorized' });
      return;
    }
  }
  next();
};

const sessions = new Map<string, SSEServerTransport>();

app.get('/sse', checkAuth, async (req: any, res: any) => {
  logger.info('SSE connection initiated');
  const transport = new SSEServerTransport('/messages', res);
  sessions.set(transport.sessionId, transport);
  res.on('close', () => {
    sessions.delete(transport.sessionId);
    logger.info(`SSE session closed: ${transport.sessionId}`);
  });
  await mcpServer.connect(transport);
});

app.post('/messages', checkAuth, async (req: any, res: any) => {
  const sessionId = req.query.sessionId as string;
  const transport = sessions.get(sessionId);
  if (!transport) {
    res.status(404).json({ error: 'Session not found' });
    return;
  }
  await transport.handlePostMessage(req, res);
});

async function warmup(): Promise<void> {
  try {
    await tokenManager.getFlowToken();
    logger.info('Flow token acquired (service.flow.microsoft.com)');
  } catch (e: any) {
    logger.warn('Flow token pre-warm failed', { error: e.message });
  }
  try {
    await tokenManager.getManagementToken();
    logger.info('Management token acquired (service.powerapps.com)');
  } catch (e: any) {
    logger.warn('Management token pre-warm failed', { error: e.message });
  }
}

app.listen(config.port, '0.0.0.0', async () => {
  logger.info(`Power Automate MCP server running on port ${config.port}`);
  logger.info(`Health: http://localhost:${config.port}/health`);
  logger.info(`SSE: http://localhost:${config.port}/sse`);
  logger.info(`Default environment: ${config.defaultEnvironmentId}`);
  await warmup();
});
