// ═══════════════════════════════════════════════════════════════
// Tool: pa-get-run-details
// Gets detailed information about a specific flow run.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const getRunDetailsSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe('The flow GUID.'),
  runId: z.string().describe('The run ID to inspect (from pa-get-run-history).'),
});

export const getRunDetailsDefinition = {
  name: 'pa-get-run-details',
  description: [
    'Gets detailed information about a specific Power Automate flow run,',
    'including full status, timing, trigger info, and error details.',
    'Use pa-get-run-history first to discover run IDs.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID.' },
      runId: { type: 'string', description: 'The specific run ID to inspect.' },
    },
    required: ['flowId', 'runId'],
  },
};

export async function executeGetRunDetails(
  args: z.infer<typeof getRunDetailsSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const run = await flowClient.getFlowRunDetails(envId, args.flowId, args.runId);
  return JSON.stringify({
    runId: run.name,
    status: run.status,
    startTime: run.startTime,
    endTime: run.endTime || null,
    triggerName: run.triggerName,
    error: run.error || null,
  }, null, 2);
}
