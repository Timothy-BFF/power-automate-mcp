// ═══════════════════════════════════════════════════════════════
// Tool: pa-trigger-flow
// Triggers a flow that has an HTTP Request trigger.
// ═══════════════════════════════════════════════════════════════

import { z } from 'zod';
import { FlowClient } from '../../clients/flow-client.js';

export const triggerFlowSchema = z.object({
  triggerUri: z.string().optional().describe(
    'The full HTTP trigger URI for the flow. Get this from pa-get-flow-details.'
  ),
  environmentId: z.string().optional().describe(
    'Power Platform environment ID (used to look up trigger URI if not provided directly).'
  ),
  flowId: z.string().optional().describe(
    'The flow GUID (used with environmentId to look up trigger URI).'
  ),
  body: z.record(z.unknown()).optional().describe(
    'JSON payload to send as the request body to the flow trigger.'
  ),
});

export const triggerFlowDefinition = {
  name: 'pa-trigger-flow',
  description: [
    'Triggers a Power Automate flow that has an HTTP Request trigger.',
    'Provide either the triggerUri directly (from pa-get-flow-details),',
    'or provide environmentId + flowId to auto-discover the trigger URI.',
    'Optionally pass a JSON body as input parameters to the flow.',
    'Returns the HTTP status code and response body from the flow execution.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {
      triggerUri: { type: 'string', description: 'Full HTTP trigger URI for the flow.' },
      environmentId: { type: 'string', description: 'Power Platform environment ID.' },
      flowId: { type: 'string', description: 'The flow GUID.' },
      body: { type: 'object', description: 'JSON payload for the flow trigger.' },
    },
  },
};

export async function executeTriggerFlow(
  args: z.infer<typeof triggerFlowSchema>,
  flowClient: FlowClient,
  defaultEnvId: string
): Promise<string> {
  let uri = args.triggerUri;

  // If no direct URI, look it up from flow details
  if (!uri) {
    const envId = args.environmentId || defaultEnvId;
    if (!envId || !args.flowId) {
      return JSON.stringify({
        error: 'Provide either triggerUri, or both environmentId and flowId.',
      });
    }
    const flow = await flowClient.getFlowDetails(envId, args.flowId);
    uri = flow.flowTriggerUri;
    if (!uri) {
      return JSON.stringify({
        error: 'This flow does not have an HTTP trigger URI. It may use a different trigger type.',
        flowId: args.flowId,
        flowName: flow.displayName,
      });
    }
  }

  const result = await flowClient.triggerFlow(uri, args.body);
  return JSON.stringify({
    triggered: true,
    statusCode: result.statusCode,
    response: result.body,
  }, null, 2);
}
