// ═══════════════════════════════════════════════════════════════
// Tool: pa-list-flows
// Lists all Power Automate flows in a specified environment.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const listFlowsSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default from config if omitted.'
  ),
  filter: z.enum(['personal', 'shared', 'all']).optional().default('all').describe(
    'Filter flows by ownership: personal (my flows), shared (team flows), or all.'
  ),
  top: z.number().optional().describe(
    'Maximum number of flows to return.'
  ),
});

export const listFlowsDefinition = {
  name: 'pa-list-flows',
  description: [
    'Lists all Power Automate flows in a Power Platform environment.',
    'Returns flow name, display name, state (Started/Stopped), created time,',
    'and last modified time. Use filter to narrow by personal or shared flows.',
    'Provide environmentId or uses the default configured environment.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID. Uses default if omitted.' },
      filter: { type: 'string', enum: ['personal', 'shared', 'all'], description: 'Filter by ownership type.' },
      top: { type: 'number', description: 'Maximum number of flows to return.' },
    },
  },
};

export async function executeListFlows(
  args: z.infer<typeof listFlowsSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const flows = await flowClient.listFlows(envId, {
    filter: args.filter,
    top: args.top,
  });

  return JSON.stringify({
    environmentId: envId,
    totalFlows: flows.length,
    flows: flows.map(f => ({
      id: f.name,
      displayName: f.displayName,
      state: f.state,
      createdTime: f.createdTime,
      lastModifiedTime: f.lastModifiedTime,
    })),
  }, null, 2);
}
