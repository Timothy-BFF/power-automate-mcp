// ═══════════════════════════════════════════════════════════════
// Tool: pa-list-environments
// Lists all Power Platform environments accessible to the service.
// ═══════════════════════════════════════════════════════════════

import { EnvironmentClient } from '../../clients/environment-client.js';

export const listEnvironmentsDefinition = {
  name: 'pa-list-environments',
  description: [
    'Lists all Power Platform environments accessible to the connected service principal.',
    'Returns environment name, display name, location, type (Production/Sandbox/Developer),',
    'state, and whether it is the default environment.',
    'Use this to discover environment IDs for other tools.',
  ].join(' '),
  inputSchema: {
    type: 'object' as const,
    properties: {},
  },
};

export async function executeListEnvironments(
  envClient: EnvironmentClient
): Promise<string> {
  const environments = await envClient.listEnvironments();

  return JSON.stringify({
    totalEnvironments: environments.length,
    environments: environments.map(e => ({
      id: e.name,
      displayName: e.displayName,
      location: e.location,
      type: e.type,
      state: e.state,
      isDefault: e.isDefault,
      createdTime: e.createdTime,
    })),
  }, null, 2);
}
