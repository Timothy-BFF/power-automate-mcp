// ═══════════════════════════════════════════════════════════════
// Tool: pa-enable-disable-flow
// Enables or disables a specific Power Automate flow.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const enableDisableFlowSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe(
    'The unique identifier (GUID) of the flow.'
  ),
  action: z.enum(['enable', 'disable']).describe(
    'Whether to enable (start) or disable (stop) the flow.'
  ),
});

export const enableDisableFlowDefinition = {
  name: 'pa-enable-disable-flow',
  description: [
    'Enables or disables a Power Automate flow.',
    'Use action "enable" to start a stopped flow,',
    'or "disable" to stop a running flow.',
    'Returns the new state after the operation.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID.' },
      action: { type: 'string', enum: ['enable', 'disable'], description: 'Enable or disable the flow.' },
    },
    required: ['flowId', 'action'],
  },
};

export async function executeEnableDisableFlow(
  args: z.infer<typeof enableDisableFlowSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const state = args.action === 'enable' ? 'Started' : 'Stopped';
  const result = await flowClient.setFlowState(envId, args.flowId, state);

  return JSON.stringify({
    flowId: args.flowId,
    action: args.action,
    success: result.success,
    newState: result.newState,
  }, null, 2);
}
