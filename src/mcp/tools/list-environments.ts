import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';

export function registerListEnvironments(server: McpServer, client: PowerPlatformClient, logger: Logger): void {
  (server as any).tool(
    'pa-list-environments',
    'Lists all Power Platform environments accessible to the configured service principal. Returns environment ID, display name, location, SKU, and lifecycle state for each environment.',
    {},
    async () => {
      try {
        const data = await client.listEnvironments();
        const envs = (data.value || []).map((env: any) => ({
          id: env.name,
          displayName: env.properties?.displayName,
          location: env.location,
          sku: env.properties?.environmentSku,
          state: env.properties?.states?.management?.id,
        }));
        return { content: [{ type: 'text', text: JSON.stringify({ totalEnvironments: envs.length, environments: envs }, null, 2) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
