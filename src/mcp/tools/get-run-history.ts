import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerGetRunHistory(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-get-run-history',
    'Gets the run history for a specific Power Automate flow. Returns run ID, status (Succeeded/Failed/Running/Cancelled), start time, end time, and trigger information.',
    {
      flowId: z.string().describe('The unique identifier of the flow.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
      top: z.number().optional().describe('Maximum number of runs to return.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const data = await client.getRunHistory(envId, args.flowId, args.top);
        const runs = (data.value || []).map((r: any) => ({
          id: r.name,
          status: r.properties?.status,
          startTime: r.properties?.startTime,
          endTime: r.properties?.endTime,
          trigger: r.properties?.trigger?.name,
        }));
        return { content: [{ type: 'text', text: JSON.stringify({ flowId: args.flowId, totalRuns: runs.length, runs }, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
