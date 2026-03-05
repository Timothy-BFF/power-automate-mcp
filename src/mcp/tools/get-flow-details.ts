import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerGetFlowDetails(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-get-flow-details',
    'Gets detailed information about a specific Power Automate flow including its definition, triggers, actions, connections, and current state.',
    {
      flowId: z.string().describe('The unique identifier of the flow.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const flow = await client.getFlowDetails(envId, args.flowId);
        const result = {
          id: flow.name,
          displayName: flow.properties?.displayName,
          state: flow.properties?.state,
          createdTime: flow.properties?.createdTime,
          lastModifiedTime: flow.properties?.lastModifiedTime,
          connectionReferences: flow.properties?.connectionReferences,
          triggers: flow.properties?.definition?.triggers,
          actions: flow.properties?.definition?.actions,
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
