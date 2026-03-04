// ═══════════════════════════════════════════════════════════════
// Tool: pa-get-flow-details
// Gets complete details about a specific flow, including its definition.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const getFlowDetailsSchema = z.object({
  environmentId: z.string().optional().describe(
    'Power Platform environment ID. Uses default if omitted.'
  ),
  flowId: z.string().describe(
    'The unique identifier (GUID) of the flow to inspect.'
  ),
});

export const getFlowDetailsDefinition = {
  name: 'pa-get-flow-details',
  description: [
    'Gets complete details about a specific Power Automate flow,',
    'including its definition (triggers, actions, conditions),',
    'state, creation time, and HTTP trigger URI if applicable.',
    'Use pa-list-flows first to discover flow IDs.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID to inspect.' },
    },
    required: ['flowId'],
  },
};

export async function executeGetFlowDetails(
  args: z.infer<typeof getFlowDetailsSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  const envId = args.environmentId || defaultEnvId;
  if (!envId) {
    return JSON.stringify({ error: 'No environmentId provided and no default configured.' });
  }

  const flow = await flowClient.getFlowDetails(envId, args.flowId);
  return JSON.stringify({
    id: flow.name,
    displayName: flow.displayName,
    state: flow.state,
    createdTime: flow.createdTime,
    lastModifiedTime: flow.lastModifiedTime,
    hasTriggerUri: !!flow.flowTriggerUri,
    triggerUri: flow.flowTriggerUri || null,
    definition: flow.definition,
  }, null, 2);
}
