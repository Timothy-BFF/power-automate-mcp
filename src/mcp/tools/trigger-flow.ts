import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerTriggerFlow(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-trigger-flow',
    'Manually triggers a Power Automate flow that has an HTTP request trigger. Optionally pass a JSON body to the trigger.',
    {
      flowId: z.string().describe('The unique identifier of the flow to trigger.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
      triggerBody: z.record(z.any()).optional().describe('Optional JSON body to pass to the flow trigger.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        const result = await client.triggerFlow(envId, args.flowId, args.triggerBody);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, flowId: args.flowId, result }) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
