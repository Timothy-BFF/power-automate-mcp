// ═══════════════════════════════════════════════════════════════
// Tool: pa-delete-flow
// Permanently deletes a Power Automate flow. Destructive operation.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const deleteFlowSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe(
    'The unique identifier (GUID) of the flow to delete.'
  ),
  confirmDelete: z.boolean().describe(
    'Must be set to true to confirm deletion. This is a destructive operation.'
  ),
});

export const deleteFlowDefinition = {
  name: 'pa-delete-flow',
  description: [
    'Permanently deletes a Power Automate flow. THIS IS DESTRUCTIVE AND IRREVERSIBLE.',
    'The confirmDelete parameter must be set to true to proceed.',
    'Always verify the flow ID with pa-get-flow-details before deleting.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID to delete.' },
      confirmDelete: { type: 'boolean', description: 'Must be true to confirm deletion.' },
    },
    required: ['flowId', 'confirmDelete'],
  },
};

export async function executeDeleteFlow(
  args: z.infer<typeof deleteFlowSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  if (!args.confirmDelete) {
    return JSON.stringify({
      error: 'Deletion not confirmed. Set confirmDelete to true to proceed.',
      warning: 'This operation is PERMANENT and IRREVERSIBLE.',
      flowId: args.flowId,
    });
  }

  const result = await flowClient.deleteFlow(envId, args.flowId);
  return JSON.stringify({
    flowId: args.flowId,
    deleted: result.success,
    message: 'Flow has been permanently deleted.',
  }, null, 2);
}
