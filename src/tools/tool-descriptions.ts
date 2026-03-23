/**
 * Tool Descriptions for Power Automate MCP
 * v3.1.0 - Added pa-get-connection and pa-delete-connection
 *
 * These descriptions are embedded in the MCP tool schema and visible
 * to AI agents when they discover available tools. The procedure
 * guidance prevents agents from making destructive decisions based
 * on misread API responses.
 */

export const TOOL_DESCRIPTIONS: Record<string, string> = {
  // =========================================================================
  // AUTH TOOLS
  // =========================================================================

  'pa-auth-start': [
    'Start Device Code Flow authentication for Power Automate write operations.',
    'MUST be called before any write operation (create, update, trigger flow).',
    'Returns a user_code and verification_uri. The user must visit the URL',
    'and enter the code to complete authentication.',
    '',
    'After calling this, poll with pa-auth-poll until status = "authenticated".',
    'DO NOT attempt write operations until authentication is confirmed.',
  ].join('\n'),

  'pa-auth-poll': [
    'Poll for Device Code Flow authentication completion.',
    'Call this after pa-auth-start to check if the user has completed login.',
    '',
    'Returns:',
    '  status: "authenticated" - user has logged in, write operations are now available',
    '  status: "pending" - user has not yet completed login, call again in 5 seconds',
    '  status: "error" - authentication failed, call pa-auth-start again',
  ].join('\n'),

  'pa-auth-status': [
    'Check current authentication status for a user.',
    'Returns whether a valid delegated token exists for write operations.',
  ].join('\n'),

  // =========================================================================
  // FLOW READ TOOLS
  // =========================================================================

  'pa-list-flows': [
    'List Power Automate flows in the environment.',
    'Uses service principal (admin) token. Returns flow names, states, and IDs.',
  ].join('\n'),

  'pa-get-flow-details': [
    'Get detailed information about a specific Power Automate flow.',
    '',
    'IMPORTANT - Read Behavior:',
    'This tool uses DUAL-PATH GET (v3.0.3):',
    '  1. If a delegated user token is available: uses the user endpoint',
    '     (returns FULL definition including triggers and actions)',
    '  2. If no user token: falls back to admin endpoint',
    '     (may return EMPTY definition due to scope limitation)',
    '',
    'CRITICAL: Check these response fields before making decisions:',
    '  - _fetchedVia: shows which path was used ("delegated" or "admin")',
    '  - _definitionStatus: "POPULATED" or "EMPTY_OR_NOT_RETURNED"',
    '  - _definitionNote: explains what to do if definition appears empty',
    '',
    'If _fetchedVia = "admin" and definition is empty:',
    '  This is a known scope limitation, NOT a missing definition.',
    '  The definition IS saved if the create/update returned 200.',
    '  DO NOT delete and recreate the flow.',
    '  Either re-authenticate (pa-auth-start) and call again,',
    '  or ask the user to verify in the Power Automate portal.',
  ].join('\n'),

  // =========================================================================
  // FLOW WRITE TOOLS
  // =========================================================================

  'pa-create-flow': [
    'Create a new Power Automate flow with a complete definition.',
    '',
    '*** REQUIRED PROCEDURE (follow this exact sequence): ***',
    '',
    'STEP 1 - AUTHENTICATE FIRST:',
    '  Call pa-auth-start, then pa-auth-poll until status = "authenticated".',
    '  DO NOT call pa-create-flow without authentication.',
    '',
    'STEP 2 - CREATE WITH FULL DEFINITION:',
    '  Pass the COMPLETE definition including triggers AND actions.',
    '  The tool sends shell + definition in ONE API call.',
    '  Do NOT create a shell first and update separately.',
    '',
    'STEP 3 - TRUST THE RESPONSE:',
    '  If the response includes status: "created" and _authType: "delegated",',
    '  the definition IS saved. Do not second-guess a 200/201 response.',
    '',
    'STEP 4 - WAIT BEFORE VERIFYING:',
    '  Wait at least 10 seconds before calling pa-get-flow-details.',
    '  NEVER call pa-get-flow-details in parallel with pa-create-flow.',
    '  Power Automate has a propagation delay between write and read.',
    '',
    'STEP 5 - VERIFY CORRECTLY:',
    '  When checking with pa-get-flow-details, always check _fetchedVia.',
    '  If _fetchedVia = "admin" and definition empty, this is a scope limitation.',
    '  The definition IS there. DO NOT delete and recreate.',
    '',
    '*** NEVER delete a flow because get-flow-details shows empty definition. ***',
  ].join('\n'),

  'pa-update-flow': [
    'Update an existing Power Automate flow (definition, name, state, connections).',
    '',
    'REQUIRES: Delegated user authentication (call pa-auth-start first).',
    '',
    'Can update: displayName, definition, state, connectionReferences.',
    'Send only the fields you want to change.',
    '',
    'If response returns 200, the update IS saved.',
    'Same verification rules as pa-create-flow apply:',
    '  - Wait 10+ seconds before verifying',
    '  - Check _fetchedVia on the get-flow-details response',
    '  - Admin path may show empty definition (scope limitation)',
  ].join('\n'),

  // =========================================================================
  // FLOW LIFECYCLE TOOLS
  // =========================================================================

  'pa-enable-flow': [
    'Enable (start) a Power Automate flow.',
    'Note: Flow connections must be authorized in the Power Automate portal',
    'before enabling. If connections are not authorized, the flow will fail.',
  ].join('\n'),

  'pa-disable-flow': [
    'Disable (stop) a Power Automate flow.',
  ].join('\n'),

  'pa-delete-flow': [
    'Delete a Power Automate flow permanently.',
    '',
    'WARNING: Only delete a flow if you are certain it should be removed.',
    'Do NOT delete a flow because pa-get-flow-details shows an empty definition.',
    'An empty definition from the admin GET path is a known scope limitation.',
    'Check _fetchedVia and _definitionNote before deciding to delete.',
  ].join('\n'),

  'pa-trigger-flow': [
    'Manually trigger a Power Automate flow.',
    '',
    'REQUIRES: Delegated user authentication and a flow with a manual/Request trigger.',
    'Recurrence-triggered flows cannot be triggered via this endpoint.',
    'They run on their configured schedule.',
  ].join('\n'),

  // =========================================================================
  // FLOW RUN TOOLS
  // =========================================================================

  'pa-get-run-history': [
    'Get the run history for a Power Automate flow.',
    'Returns recent runs with status, start/end times, and trigger info.',
  ].join('\n'),

  'pa-get-run-details': [
    'Get details for a specific flow run.',
    'Returns the run status, duration, trigger, and action results.',
  ].join('\n'),

  'pa-cancel-run': [
    'Cancel a running flow execution.',
  ].join('\n'),

  // =========================================================================
  // ENVIRONMENT & CONNECTION TOOLS
  // =========================================================================

  'pa-list-environments': [
    'List all Power Platform environments accessible to the service principal.',
  ].join('\n'),

  'pa-list-connections': [
    'List API connections in the environment.',
    'Shows connection status, type, and authorization state.',
  ].join('\n'),

  'pa-get-connection': [
    'Get details for a specific API connection.',
    'Returns connection status, connector type, authorization state, and creator info.',
    'Requires the connectionId (name field) from pa-list-connections.',
  ].join('\n'),

  'pa-delete-connection': [
    'Delete an API connection from the environment permanently.',
    '',
    'WARNING: Deleting a connection may break flows that depend on it.',
    'Check which flows reference this connection before deleting.',
    'This action cannot be undone.',
  ].join('\n'),
};

/**
 * Flow Creation Procedure Summary (for embedding in system prompts)
 *
 * This is a compact version of the full procedure from
 * docs/skills/flow-creation-procedure.md
 */
export const FLOW_CREATION_PROCEDURE = [
  'FLOW CREATION SEQUENCE (mandatory):',
  '1. pa-auth-start -> pa-auth-poll (wait for "authenticated")',
  '2. pa-create-flow with FULL definition (triggers + actions in one call)',
  '3. Wait 10+ seconds (propagation delay)',
  '4. pa-get-flow-details -> check _fetchedVia and _definitionStatus',
  '5. If _fetchedVia="admin" and empty: scope limitation, NOT missing definition',
  '',
  'NEVER delete a flow because get-flow-details shows empty definition.',
  'NEVER call get-flow-details in parallel with create-flow.',
  'If create returned 200 with _authType="delegated", the definition IS saved.',
].join('\n');
