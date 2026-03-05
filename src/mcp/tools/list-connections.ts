import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerListConnections(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-list-connections',
    'Lists all Power Platform connections in an environment. Returns connection ID, display name, status, connector information, and creation time.',
    {
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const data = await client.listConnections(envId);
        const connections = (data.value || []).map((c: any) => ({
          id: c.name,
          displayName: c.properties?.displayName,
          status: c.properties?.statuses?.[0]?.status,
          connectorName: c.properties?.apiId,
          createdTime: c.properties?.createdTime,
        }));
        return { content: [{ type: 'text', text: JSON.stringify({ environmentId: envId, totalConnections: connections.length, connections }, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
