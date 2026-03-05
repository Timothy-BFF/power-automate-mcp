import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerGetRunDetails(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-get-run-details',
    'Gets detailed information about a specific flow run including action-level results, inputs, outputs, and timing.',
    {
      flowId: z.string().describe('The unique identifier of the flow.'),
      runId: z.string().describe('The unique identifier of the run.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const run = await client.getRunDetails(envId, args.flowId, args.runId);
        return { content: [{ type: 'text', text: JSON.stringify(run, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
