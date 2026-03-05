import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerCancelRun(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-cancel-run',
    'Cancels a currently running Power Automate flow run.',
    {
      flowId: z.string().describe('The unique identifier of the flow.'),
      runId: z.string().describe('The unique identifier of the run to cancel.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        await client.cancelRun(envId, args.flowId, args.runId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, flowId: args.flowId, runId: args.runId, message: 'Run cancelled successfully.' }) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
