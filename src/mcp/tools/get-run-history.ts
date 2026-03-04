// ═══════════════════════════════════════════════════════════════
// Tool: pa-get-run-history
// Gets the execution history for a specific flow.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const getRunHistorySchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe(
    'The unique identifier (GUID) of the flow.'
  ),
  top: z.number().optional().default(10).describe(
    'Maximum number of runs to return. Defaults to 10.'
  ),
  status: z.enum(['Succeeded', 'Failed', 'Running', 'Cancelled']).optional().describe(
    'Filter runs by status.'
  ),
});

export const getRunHistoryDefinition = {
  name: 'pa-get-run-history',
  description: [
    'Gets the execution run history for a specific Power Automate flow.',
    'Returns run status (Succeeded, Failed, Running, Cancelled),',
    'start time, end time, trigger name, and error details for failed runs.',
    'Use top to limit results and status to filter by outcome.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID.' },
      top: { type: 'number', description: 'Max runs to return (default 10).' },
      status: { type: 'string', enum: ['Succeeded', 'Failed', 'Running', 'Cancelled'], description: 'Filter by run status.' },
    },
    required: ['flowId'],
  },
};

export async function executeGetRunHistory(
  args: z.infer<typeof getRunHistorySchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const runs = await flowClient.getFlowRuns(envId, args.flowId, {
    top: args.top,
    status: args.status,
  });

  return JSON.stringify({
    flowId: args.flowId,
    totalRuns: runs.length,
    runs: runs.map(r => ({
      runId: r.name,
      status: r.status,
      startTime: r.startTime,
      endTime: r.endTime || null,
      triggerName: r.triggerName,
      error: r.error || null,
    })),
  }, null, 2);
}
