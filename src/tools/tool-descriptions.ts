/**
 * Tool Descriptions — Power Automate MCP
 * v3.3.0: 26 total tools, unified registration
 *
 * IMPORTANT: This must be a Record<string, string> — NOT an array.
 * index.ts accesses it as TOOL_DESCRIPTIONS['pa-xxx'].
 *
 * v3.2.0 → v3.3.0 changes:
 *   FIXED: pa-list-connections (now uses PowerApps admin host)
 *   NEW: pa-create-solution, pa-delete-solution, pa-create-connection
 */

export const TOOL_DESCRIPTIONS: Record<string, string> = {
  // ---- Flow Management (8 tools) ----
  'pa-list-flows': 'List all flows in a Power Platform environment. Returns flow names, IDs, states, and trigger information.',
  'pa-get-flow-details': 'Get detailed information about a specific flow, including its full definition, triggers, actions, and connections.',
  'pa-create-flow': 'Create a new Power Automate flow. Provide the FULL flow definition in one call. Requires user authentication (pa-auth-start).',
  'pa-update-flow': 'Update an existing Power Automate flow. Can update display name, definition, state, or connection references.',
  'pa-enable-flow': 'Enable (turn on) a Power Automate flow.',
  'pa-disable-flow': 'Disable (turn off) a Power Automate flow.',
  'pa-delete-flow': 'Delete a Power Automate flow permanently. This cannot be undone.',
  'pa-trigger-flow': 'Manually trigger a Power Automate flow. Works with flows that have manual/HTTP triggers.',

  // ---- Run Management (3 tools) ----
  'pa-get-run-history': 'Get the run history for a specific flow. Returns recent runs with status, start/end times, and trigger info.',
  'pa-get-run-details': 'Get detailed information about a specific flow run, including action results and error details.',
  'pa-cancel-run': 'Cancel a running flow instance.',

  // ---- Environment Management (1 tool) ----
  'pa-list-environments': 'List all Power Platform environments accessible to the service principal.',

  // ---- Connection Management (4 tools — 3 existing + 1 new) ----
  'pa-list-connections': 'List all connections in a Power Platform environment. Uses PowerApps admin API (no user auth required).',
  'pa-get-connection': 'Get detailed information about a specific connection, including status, connector type, and creation details.',
  'pa-delete-connection': 'Delete a connection permanently.',
  'pa-create-connection': 'Create a new connection to a connector in Power Automate. Requires user authentication (pa-auth-start). The connection will be owned by the authenticated user.',

  // ---- Solution Management (7 tools — 5 existing + 2 new) ----
  'pa-list-solutions': 'List Dataverse solutions. By default shows only unmanaged solutions.',
  'pa-get-solution': 'Get detailed information about a specific Dataverse solution by unique name or GUID.',
  'pa-list-solution-components': 'List all components (flows, entities, web resources, etc.) within a Dataverse solution.',
  'pa-export-solution': 'Export a Dataverse solution as a ZIP file (base64 encoded).',
  'pa-add-solution-component': 'Add an existing component (flow, entity, etc.) to a Dataverse solution.',
  'pa-create-solution': 'Create a new Dataverse solution. Requires a publisher ID (use pa-get-solution on an existing solution to find valid publisher IDs).',
  'pa-delete-solution': 'Delete a Dataverse solution. WARNING: This permanently removes the solution container. Components inside may remain in the environment.',

  // ---- Auth Tools (3 tools) ----
  'pa-auth-start': 'Start OAuth device code flow. Returns a user code and URL for sign-in at microsoft.com/devicelogin.',
  'pa-auth-poll': 'Poll for authentication completion after pa-auth-start. Returns authenticated, pending, or expired.',
  'pa-auth-status': 'Check current authentication status for a user or all users.',
};
