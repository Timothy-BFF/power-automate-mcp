import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../api/power-platform-client.js';
import { registerListEnvironments } from './tools/list-environments.js';
import { registerListFlows } from './tools/list-flows.js';
import { registerGetFlowDetails } from './tools/get-flow-details.js';
import { registerGetRunHistory } from './tools/get-run-history.js';
import { registerGetRunDetails } from './tools/get-run-details.js';
import { registerListConnections } from './tools/list-connections.js';
import { registerEnableDisableFlow } from './tools/enable-disable-flow.js';
import { registerDeleteFlow } from './tools/delete-flow.js';
import { registerTriggerFlow } from './tools/trigger-flow.js';
import { registerCancelRun } from './tools/cancel-run.js';

export function createAndConfigureMcpServer(
  client: PowerPlatformClient,
  defaultEnvId: string,
  logger: Logger
): McpServer {
  const server = new McpServer({ name: 'power-automate-mcp', version: '2.0.0' });

  registerListEnvironments(server, client, logger);
  registerListFlows(server, client, defaultEnvId, logger);
  registerGetFlowDetails(server, client, defaultEnvId, logger);
  registerGetRunHistory(server, client, defaultEnvId, logger);
  registerGetRunDetails(server, client, defaultEnvId, logger);
  registerListConnections(server, client, defaultEnvId, logger);
  registerEnableDisableFlow(server, client, defaultEnvId, logger);
  registerDeleteFlow(server, client, defaultEnvId, logger);
  registerTriggerFlow(server, client, defaultEnvId, logger);
  registerCancelRun(server, client, defaultEnvId, logger);

  logger.info('MCP tools registered: 10');
  return server;
}
