import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerDeleteFlow(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-delete-flow',
    'Permanently deletes a Power Automate flow. This action cannot be undone. Use with caution.',
    {
      flowId: z.string().describe('The unique identifier of the flow to delete.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        await client.deleteFlow(envId, args.flowId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, flowId: args.flowId, message: 'Flow deleted permanently.' }) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
