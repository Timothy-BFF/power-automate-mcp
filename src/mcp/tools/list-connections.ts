// ═══════════════════════════════════════════════════════════════
// Tool: pa-list-connections
// Lists all API connections in a Power Platform environment.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { ConnectionClient } from '../../clients/connection-client.js';

export const listConnectionsSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
});

export const listConnectionsDefinition = {
  name: 'pa-list-connections',
  description: [
    'Lists all API connections configured in a Power Platform environment.',
    'Returns connection name, connector type (e.g., Office 365, SharePoint, SQL),',
    'status, and creation time. Useful for auditing or debugging flow dependencies.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
    },
  },
};

export async function executeListConnections(
  args: z.infer<typeof listConnectionsSchema>,
  connectionClient: ConnectionClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const connections = await connectionClient.listConnections(envId);

  return JSON.stringify({
    environmentId: envId,
    totalConnections: connections.length,
    connections: connections.map(c => ({
      id: c.name,
      displayName: c.displayName,
      connectorName: c.connectorName,
      status: c.status,
      createdTime: c.createdTime,
    })),
  }, null, 2);
}
