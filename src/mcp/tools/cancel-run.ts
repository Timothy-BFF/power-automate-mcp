// ═══════════════════════════════════════════════════════════════
// Tool: pa-cancel-run
// Cancels a currently running flow execution.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const cancelRunSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe('The flow GUID.'),
  runId: z.string().describe('The run ID to cancel (from pa-get-run-history).'),
});

export const cancelRunDefinition = {
  name: 'pa-cancel-run',
  description: [
    'Cancels a currently running Power Automate flow execution.',
    'Only works on runs that are in "Running" status.',
    'Use pa-get-run-history with status filter "Running" to find active runs.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID.' },
      runId: { type: 'string', description: 'The run ID to cancel.' },
    },
    required: ['flowId', 'runId'],
  },
};

export async function executeCancelRun(
  args: z.infer<typeof cancelRunSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const result = await flowClient.cancelFlowRun(envId, args.flowId, args.runId);
  return JSON.stringify({
    flowId: args.flowId,
    runId: args.runId,
    cancelled: result.success,
  }, null, 2);
}
