/**
 * REST Transport Skill Handlers — Power Automate MCP v3.3.0
 *
 * The MCP SDK natively handles prompts/resources for SSE transport.
 * This module provides equivalent support for the REST JSON-RPC transport.
 *
 * Handles: prompts/list, prompts/get, resources/list, resources/read
 */

const PROMPT_REGISTRY = [
  {
    name: 'workflow-auth',
    description: 'Step-by-step guide for authenticating a user via Device Code Flow. Read this BEFORE calling pa-auth-start.',
    arguments: [
      { name: 'user_id', description: 'User email address (e.g., jose@company.com)', required: false }
    ]
  },
  {
    name: 'workflow-create-flow',
    description: 'MANDATORY 5-step procedure for creating a Power Automate flow. Includes definition JSON format requirements and solution placement.',
    arguments: [
      { name: 'displayName', description: 'Name of the flow to create', required: false },
      { name: 'user_id', description: 'Authenticated user email', required: false }
    ]
  },
  {
    name: 'workflow-create-connection',
    description: 'Complete guide for creating a connector connection, including expected OAuth consent flow.',
    arguments: [
      { name: 'connectorId', description: 'Connector name (e.g., shared_office365)', required: false },
      { name: 'user_id', description: 'Authenticated user email', required: false }
    ]
  }
];

const RESOURCE_REGISTRY = [
  {
    uri: 'power-automate://docs/parameter-conventions',
    name: 'parameter-conventions',
    mimeType: 'text/markdown',
    description: 'camelCase vs snake_case mapping for all 26 tools and definition body properties'
  },
  {
    uri: 'power-automate://docs/connection-lifecycle',
    name: 'connection-lifecycle',
    mimeType: 'text/markdown',
    description: 'OAuth vs non-OAuth connector behavior, error codes, and consent flow'
  },
  {
    uri: 'power-automate://config/environment',
    name: 'environment-info',
    mimeType: 'text/markdown',
    description: 'Current environment ID, Dataverse URL, auth architecture (dynamic)'
  }
];

function generateWorkflowAuth(args: Record<string, string> = {}): any {
  const userId = args.user_id || '<user-email@company.com>';
  return {
    messages: [{
      role: 'user',
      content: {
        type: 'text',
        text: [
          '# Authentication Workflow — Device Code Flow',
          '',
          '## When to Use:',
          'Any write operation (create flow, create connection, update flow) requires user authentication.',
          'Read operations (list flows, list environments, get details) use the service principal — no user auth needed.',
          '',
          '## Steps:',
          `1. **Start auth**: Call \`pa-auth-start\` with user_id: "${userId}"`,
          '2. **Show code**: Tell the user: "Please go to https://microsoft.com/devicelogin and enter code: {user_code}"',
          '3. **Wait**: Give the user 30–60 seconds to complete sign-in',
          '4. **Poll**: Call `pa-auth-poll` with the same user_id',
          '   - "pending" → Wait 15–30 seconds and poll again',
          '   - "authenticated" → Success! Proceed with write operations',
          '   - "error" → Start over with pa-auth-start',
          '',
          '## Token Lifetime:',
          '- Access token: ~75 min (auto-refreshed). Refresh token: ~90 days (auto-rotated)',
          '',
          '## IMPORTANT — Environment Discovery:',
          '- pa-list-environments may return empty — this is a known BAP admin role limitation',
          `- The production environment ID is already configured: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not set)'}`,
          '- All tools default to the production environment when environmentId is omitted',
          '- Just call pa-list-flows, pa-create-flow, etc. directly — no environmentId needed',
          '',
          '## Parameter Format:',
          '- user_id: Full email address (e.g., "jose.m.sanchez@bolthousefresh.com")',
        ].join('\n')
      }
    }]
  };
}

function generateWorkflowCreateFlow(args: Record<string, string> = {}): any {
  const displayName = args.displayName || '<provide-display-name>';
  const userId = args.user_id || '<must-be-authenticated>';
  return {
    messages: [{
      role: 'user',
      content: {
        type: 'text',
        text: [
          '# Flow Creation — Mandatory 5-Step Procedure',
          '',
          `**Flow**: ${displayName}`,
          `**User**: ${userId}`,
          '',
          '## Steps (MUST follow in order):',
          '',
          '### Step 1: Authenticate',
          '- Call pa-auth-start with user_id → receive device code',
          '- User signs in at microsoft.com/devicelogin',
          '- Call pa-auth-poll → wait for status "authenticated"',
          '',
          '### Step 2: Create Flow (SINGLE CALL — FULL DEFINITION)',
          '- Call pa-create-flow with the COMPLETE definition in ONE call',
          '- Required: displayName + definition',
          '- Optional: connectionReferences, state ("Stopped" default), environmentId',
          '',
          '### ⚠️ Flow Definition JSON Format (CRITICAL):',
          '',
          'The definition MUST include these required top-level properties:',
          '  {',
          '    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",',
          '    "contentVersion": "1.0.0.0",',
          '    "triggers": { ... },',
          '    "actions": { ... }',
          '  }',
          '',
          'Properties inside actions MUST use camelCase — NOT snake_case:',
          '  ✅ runAfter          ❌ run_after',
          '  ✅ defaultValue      ❌ default_value',
          '  ✅ triggerConditions  ❌ trigger_conditions',
          '  ✅ operationId       ❌ operation_id',
          '  ✅ contentType       ❌ content_type',
          '  ✅ retryPolicy       ❌ retry_policy',
          '',
          '### ⚠️ $connections Parameter Format (CRITICAL):',
          '',
          'If the flow uses connectors, include a parameters block with $connections:',
          '  "parameters": {',
          '    "$connections": {',
          '      "defaultValue": {},',
          '      "type": "Object"',
          '    }',
          '  }',
          '',
          'Use "defaultValue" NOT "default_value". The server auto-normalizes as a safety net,',
          'but correct camelCase avoids API round-trip errors.',
          '',
          '### Step 3: Wait for Propagation',
          '- Wait at least 10 seconds before verifying',
          '',
          '### Step 4: Verify',
          '- Call pa-get-flow-details with the returned flowId',
          '- Check _fetchedVia, _definitionStatus, _authType',
          '',
          '### Step 5: Add to Solution (if applicable)',
          '',
          'If the flow should live inside a solution:',
          '- Call `pa-add-solution-component` IMMEDIATELY after Step 4:',
          '  - solutionUniqueName: The target solution unique name (e.g., "BeakSolution")',
          '  - componentId: The flow ID returned from pa-create-flow in Step 2',
          '  - componentType: 29 (Cloud Flow)',
          '',
          '⚠️ You MUST use the flow ID from pa-create-flow (Step 2).',
          'Do NOT try to "find" the flow via pa-list-flows first — flows created via',
          'the API are immediately available by their ID but may take 15–30 minutes',
          'to appear in list results due to admin API indexing delays.',
          '',
          '⚠️ Flows created via the Power Automate portal UI (NOT via pa-create-flow)',
          'may not appear in pa-list-flows for 15–30 minutes. To add a UI-created flow',
          'to a solution, get the flow GUID from the browser URL:',
          '  .../environments/{envId}/flows/{FLOW-GUID}/details',
          'Then call pa-add-solution-component with that GUID as componentId.',
          '',
          '## ⚠️ CRITICAL RULES:',
          '',
          '### NEVER Delete on Empty Definition',
          'If pa-get-flow-details shows empty definition but _fetchedVia was "admin"',
          'and _authType was "delegated" → the flow WAS created correctly.',
          'The admin endpoint does NOT return definitions. This is expected.',
          '',
          '### ALWAYS Full Definition in One Call',
          'Do NOT create an empty flow and PATCH the definition later.',
          '',
          '### Parameter Format',
          'ALL tool parameters use camelCase: flowId, displayName, environmentId, connectionReferences.',
          'ALL definition body properties use camelCase: runAfter, defaultValue, operationId.',
        ].join('\n')
      }
    }]
  };
}

function generateWorkflowCreateConnection(args: Record<string, string> = {}): any {
  const connectorId = args.connectorId || '<provide-connector-id>';
  const userId = args.user_id || '<must-be-authenticated>';
  return {
    messages: [{
      role: 'user',
      content: {
        type: 'text',
        text: [
          '# Connection Creation Workflow',
          '',
          `**Connector**: ${connectorId}`,
          `**User**: ${userId}`,
          '',
          '## Pre-requisite:',
          'User MUST be authenticated first (pa-auth-start → pa-auth-poll → "authenticated").',
          '',
          '## Steps:',
          `1. Call pa-create-connection with connectorId: "${connectorId}", user_id: "${userId}"`,
          '2. environmentId: Optional — uses production environment if omitted',
          '3. connectionParameters: Optional — only for SQL, HTTP, custom connectors',
          '',
          '## ⚠️ EXPECTED BEHAVIOR — OAuth Connectors:',
          'For Office 365, SharePoint, Teams, Outlook, OneDrive:',
          '- Response WILL show: connectionStatus: "Error", needsConsent: true',
          '- This is 100% NORMAL — not a bug, not a failure!',
          '- Tell user: "Go to make.powerautomate.com → Data → Connections → Authorize"',
          '',
          '## Non-OAuth Connectors (HTTP, SQL):',
          'Return connectionStatus: "Connected" immediately. No manual step needed.',
          '',
          '## Common Connector IDs:',
          '| Connector | connectorId | Manual Auth? |',
          '|-----------|------------|-------------|',
          '| Office 365 | shared_office365 | Yes |',
          '| SharePoint | shared_sharepointonline | Yes |',
          '| Teams | shared_teams | Yes |',
          '| OneDrive | shared_onedrive | Yes |',
          '| Dataverse | shared_commondataserviceforapps | Yes |',
          '| Approvals | shared_approvals | No |',
          '| HTTP | shared_webcontents | No |',
          '',
          '## PARAMETER FORMAT:',
          'Always use camelCase: connectorId, environmentId, connectionParameters',
        ].join('\n')
      }
    }]
  };
}

const PROMPT_GENERATORS: Record<string, (args: Record<string, string>) => any> = {
  'workflow-auth': generateWorkflowAuth,
  'workflow-create-flow': generateWorkflowCreateFlow,
  'workflow-create-connection': generateWorkflowCreateConnection,
};

function generateParameterConventions(uri: string): any {
  return {
    contents: [{
      uri,
      mimeType: 'text/markdown',
      text: [
        '# Parameter Naming Conventions — Power Automate MCP',
        '',
        '## All tool parameters use camelCase:',
        '| Correct (camelCase) | Wrong (snake_case) |',
        '|---------------------|-------------------|',
        '| connectorId | connector_id |',
        '| environmentId | environment_id |',
        '| flowId | flow_id |',
        '| connectionId | connection_id |',
        '| solutionId | solution_id |',
        '| displayName | display_name |',
        '| connectionReferences | connection_references |',
        '| connectionParameters | connection_parameters |',
        '| solutionUniqueName | solution_unique_name |',
        '| componentId | component_id |',
        '| friendlyName | friendly_name |',
        '| uniqueName | unique_name |',
        '| publisherId | publisher_id |',
        '',
        '## Exception: user_id uses snake_case (auth tools)',
        '',
        '## IMPORTANT — Definition Body Properties:',
        'Properties INSIDE flow definitions also use camelCase:',
        '| Correct | Wrong |',
        '|---------|-------|',
        '| runAfter | run_after |',
        '| defaultValue | default_value |',
        '| triggerConditions | trigger_conditions |',
        '| operationId | operation_id |',
        '| contentType | content_type |',
        '| retryPolicy | retry_policy |',
        '',
        '## Auto-Resolution:',
        'The server auto-maps snake_case to camelCase for both tool params AND definition properties.',
      ].join('\n'),
    }],
  };
}

function generateConnectionLifecycle(uri: string): any {
  return {
    contents: [{
      uri,
      mimeType: 'text/markdown',
      text: [
        '# Connection Lifecycle Guide',
        '',
        '## OAuth Connectors (Office 365, SharePoint, Teams, Outlook, OneDrive):',
        'pa-create-connection creates a "shell" connection:',
        '  Status: Error, needsConsent: true',
        '  User authorizes at make.powerautomate.com → Data → Connections',
        '  Status changes to: Connected',
        '',
        'IMPORTANT: "Error" + needsConsent is EXPECTED BEHAVIOR, not a failure!',
        '',
        '## Non-OAuth Connectors (HTTP, SQL):',
        'Returns "Connected" immediately. No manual step needed.',
        '',
        '## Error Codes:',
        '- 404: Connector name wrong or does not exist',
        '- 403: Admin consent for PowerApps Service needed in Azure AD',
        '- 409: Connection name already exists',
        '- 401: User token expired — re-authenticate with pa-auth-start',
      ].join('\n'),
    }],
  };
}

function generateEnvironmentInfo(uri: string): any {
  return {
    contents: [{
      uri,
      mimeType: 'text/markdown',
      text: [
        '# Current Environment Configuration',
        '',
        `- Environment ID: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not configured)'}`,
        `- Dataverse URL: ${process.env.DATAVERSE_URL || '(not configured)'}`,
        `- Tenant: ${process.env.AZURE_TENANT_ID ? 'Configured' : 'Not configured'}`,
        '',
        '## IMPORTANT — Environment Discovery:',
        '- pa-list-environments may return EMPTY (BAP admin role limitation)',
        `- The environment ID is: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not set)'}`,
        '- All tools default to this environment when environmentId is omitted',
        '- Just call tools directly without environmentId',
        '',
        '## Auth: Service Principal (read) + User Delegated (write)',
        '## Capabilities: 26 tools + 3 prompts + 3 resources',
      ].join('\n'),
    }],
  };
}

const RESOURCE_GENERATORS: Record<string, (uri: string) => any> = {
  'power-automate://docs/parameter-conventions': generateParameterConventions,
  'power-automate://docs/connection-lifecycle': generateConnectionLifecycle,
  'power-automate://config/environment': generateEnvironmentInfo,
};

export function processSkillRequest(method: string, params: any, id: any): any | null {
  switch (method) {
    case 'prompts/list':
      console.log('[REST skills] prompts/list');
      return { jsonrpc: '2.0', id, result: { prompts: PROMPT_REGISTRY } };

    case 'prompts/get': {
      const promptName = params?.name;
      console.log(`[REST skills] prompts/get: ${promptName}`);
      const generator = PROMPT_GENERATORS[promptName];
      if (!generator) {
        return { jsonrpc: '2.0', id, error: { code: -32602, message: `Unknown prompt: ${promptName}. Available: ${Object.keys(PROMPT_GENERATORS).join(', ')}` } };
      }
      return { jsonrpc: '2.0', id, result: generator(params?.arguments || {}) };
    }

    case 'resources/list':
      console.log('[REST skills] resources/list');
      return { jsonrpc: '2.0', id, result: { resources: RESOURCE_REGISTRY } };

    case 'resources/read': {
      const uri = params?.uri;
      console.log(`[REST skills] resources/read: ${uri}`);
      const generator = RESOURCE_GENERATORS[uri];
      if (!generator) {
        return { jsonrpc: '2.0', id, error: { code: -32602, message: `Unknown resource: ${uri}. Available: ${Object.keys(RESOURCE_GENERATORS).join(', ')}` } };
      }
      return { jsonrpc: '2.0', id, result: generator(uri) };
    }

    default:
      return null;
  }
}
