import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerListFlows(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-list-flows',
    'Lists all Power Automate flows in a Power Platform environment. Returns flow name, display name, state (Started/Stopped), created time, and last modified time. Use filter to narrow by personal or shared flows. Provide environmentId or uses the default configured environment.',
    {
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
      filter: z.enum(['personal', 'shared', 'all']).optional().describe('Filter by ownership type: personal (my flows), shared (team flows), or all.'),
      top: z.number().optional().describe('Maximum number of flows to return.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const data = await client.listFlows(envId, args.filter, args.top);
        const flows = (data.value || []).map((f: any) => ({
          id: f.name,
          displayName: f.properties?.displayName,
          state: f.properties?.state,
          createdTime: f.properties?.createdTime,
          lastModifiedTime: f.properties?.lastModifiedTime,
        }));
        return { content: [{ type: 'text', text: JSON.stringify({ environmentId: envId, totalFlows: flows.length, flows }, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
