import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Logger } from 'winston';
import { PowerPlatformClient } from '../../api/power-platform-client.js';
import { resolveEnvironmentId } from '../../config/environment-resolver.js';

export function registerEnableDisableFlow(server: McpServer, client: PowerPlatformClient, defaultEnvId: string, logger: Logger): void {
  (server as any).tool(
    'pa-enable-disable-flow',
    'Enables or disables a Power Automate flow. Use action "start" to enable or "stop" to disable the flow.',
    {
      flowId: z.string().describe('The unique identifier of the flow.'),
      action: z.enum(['start', 'stop']).describe('Action to perform: start (enable) or stop (disable) the flow.'),
      environmentId: z.string().optional().describe('Power Platform environment ID. Uses default if omitted.'),
    },
    async (args: any) => {
      try {
        const envId = resolveEnvironmentId(args.environmentId);
        if (args.action === 'start') {
          await client.enableFlow(envId, args.flowId);
        } else {
          await client.disableFlow(envId, args.flowId);
        }
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, flowId: args.flowId, action: args.action, message: `Flow ${args.action === 'start' ? 'enabled' : 'disabled'} successfully.` }) }] };
      } catch (err: any) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }] };
      }
    }
  );
}
