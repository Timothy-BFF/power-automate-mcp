/**
 * REST Transport Skill Handlers — Power Automate MCP v3.3.0
 *
 * The MCP SDK natively handles prompts/resources for SSE transport.
 * This module provides equivalent support for the REST JSON-RPC transport,
 * so agents like Jose's Claude can access all 6 skills via REST.
 *
 * Handles: prompts/list, prompts/get, resources/list, resources/read
 *
 * Usage in index.ts:
 *   import { processSkillRequest } from './skills/rest-skills.js';
 *
 *   // In processJsonRpcRequest, before the switch(method):
 *   const skillResult = processSkillRequest(method, params, id);
 *   if (skillResult) return skillResult;
 */

// =============================================================================
// Prompt Registry (metadata for prompts/list)
// =============================================================================
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
    description: 'MANDATORY procedure for creating a Power Automate flow. Agents MUST follow this exact sequence to avoid data loss.',
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

// =============================================================================
// Resource Registry (metadata for resources/list)
// =============================================================================
const RESOURCE_REGISTRY = [
  {
    uri: 'power-automate://docs/parameter-conventions',
    name: 'parameter-conventions',
    mimeType: 'text/markdown',
    description: 'camelCase vs snake_case mapping for all 26 tools'
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

// =============================================================================
// Prompt Content Generators (for prompts/get)
// Replicates src/skills/prompts.ts content for REST transport
// =============================================================================
function generateWorkflowAuth(args: Record<string, string> = {}): any {
  const userId = args.user_id || '<user-email@company.com>';
  return {
    messages: [{
      role: 'user',
      content: {
        type: 'text',
        text: [
          '# Authentication Workflow \u2014 Device Code Flow',
          '',
          '## When to Use:',
          'Any write operation (create flow, create connection, update flow) requires user authentication.',
          'Read operations (list flows, list environments, get details) use the service principal \u2014 no user auth needed.',
          '',
          '## Steps:',
          `1. **Start auth**: Call \`pa-auth-start\` with user_id: "${userId}"`,
          '2. **Show code**: Tell the user exactly:',
          '   "Please go to https://microsoft.com/devicelogin and enter code: {user_code}"',
          '3. **Wait**: Give the user 30\u201360 seconds to complete sign-in in their browser',
          '4. **Poll**: Call `pa-auth-poll` with the same user_id',
          '   - Status **"pending"** \u2192 User hasn\'t completed sign-in yet. Wait 15\u201330 seconds and poll again.',
          '   - Status **"authenticated"** \u2192 Success! Proceed with write operations.',
          '   - Status **"error"** \u2192 Auth failed. Start over with pa-auth-start.',
          '',
          '## Token Lifetime:',
          '- Access token: ~75 minutes (auto-refreshed silently by the server)',
          '- Refresh token: ~90 days (rotated automatically on each use)',
          '- No re-auth prompt needed unless the refresh token expires or the server redeploys',
          '',
          '## After Server Redeployment:',
          '- Users MUST open a FRESH SSE connection',
          '- Users MUST re-authenticate (refresh tokens are stored in-memory only)',
          '',
          '## Multi-Scope Token Exchange:',
          'The server automatically exchanges the Flow-scoped token for PowerApps-scoped tokens',
          'when needed (e.g., for pa-create-connection). Users only authenticate once.',
          '',
          '## IMPORTANT \u2014 Environment Discovery:',
          '- `pa-list-environments` may return empty if the service principal lacks Power Platform Admin role',
          `- The production environment ID is already configured server-side: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not set)'}`,
          '- You do NOT need to discover it \u2014 all tools default to the production environment automatically',
          '- Just call pa-list-flows, pa-create-flow, etc. without passing environmentId',
          '',
          '## Parameter Format:',
          '- user_id: Full email address (e.g., "Timothy.Escamilla@bolthousefresh.com")',
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
          '# Flow Creation \u2014 Mandatory Procedure',
          '',
          `**Flow**: ${displayName}`,
          `**User**: ${userId}`,
          '',
          '## Steps (MUST follow in order):',
          '',
          '### Step 1: Authenticate',
          '- Call `pa-auth-start` with user_id \u2192 receive device code',
          '- User signs in at microsoft.com/devicelogin',
          '- Call `pa-auth-poll` \u2192 wait for status "authenticated"',
          '- See prompt: workflow-auth for full details',
          '',
          '### Step 2: Create Flow (SINGLE CALL \u2014 FULL DEFINITION)',
          '- Call `pa-create-flow` with the COMPLETE definition in ONE call',
          '- Required parameters:',
          '  - `displayName`: Human-readable flow name',
          '  - `definition`: Complete workflow JSON with triggers AND actions',
          '- Optional parameters:',
          '  - `connectionReferences`: Map of connection references for connectors used',
          '  - `state`: "Started" or "Stopped" (defaults to "Stopped" \u2014 safe for testing)',
          '  - `environmentId`: Uses production environment if omitted',
          '',
          '### Step 3: Wait for Propagation',
          '- Wait **at least 10 seconds** before verifying',
          '- The Flow service needs time to persist the definition across replicas',
          '',
          '### Step 4: Verify',
          '- Call `pa-get-flow-details` with the returned flowId',
          '- Check these metadata fields:',
          '  - `_fetchedVia`: Shows which endpoint retrieved the definition',
          '  - `_definitionStatus`: Should indicate definition is present',
          '  - `_authType`: Shows what token type was used during creation',
          '',
          '## \u26a0\ufe0f CRITICAL RULES:',
          '',
          '### NEVER Delete on Empty Definition',
          'If `pa-get-flow-details` shows an empty/missing definition but:',
          '- `_fetchedVia` was "admin" \u2192 The admin endpoint does NOT return definitions',
          '- `_authType` was "delegated" during creation \u2192 The flow WAS created correctly',
          'This is a **known API limitation**, NOT a creation failure.',
          '',
          '### ALWAYS Full Definition in One Call',
          'Do NOT create an empty flow and PATCH the definition later.',
          'This causes orphaned flows with no triggers or actions.',
          '',
          '### Parameter Format',
          'ALL parameters use camelCase: flowId, displayName, environmentId, connectionReferences.',
          '(See resource: parameter-conventions for the complete list)',
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
          'User MUST be authenticated first (pa-auth-start \u2192 pa-auth-poll \u2192 "authenticated").',
          '',
          '## Steps:',
          '1. **Create connection**: Call `pa-create-connection` with:',
          `   - \`connectorId\`: "${connectorId}" (exact connector short name)`,
          `   - \`user_id\`: "${userId}"`,
          '   - `environmentId`: Optional \u2014 uses production environment if omitted',
          '   - `connectionParameters`: Optional \u2014 only needed for SQL, HTTP, custom connectors',
          '',
          '## \u26a0\ufe0f EXPECTED BEHAVIOR \u2014 OAuth Connectors:',
          '',
          'For **Office 365, SharePoint, Teams, Outlook, OneDrive**:',
          '- The API response WILL show: `connectionStatus: "Error"`, `needsConsent: true`',
          '- **This is 100% NORMAL \u2014 not a bug, not a failure!**',
          '- The API creates a "shell" connection requiring interactive browser authorization',
          '',
          '**Tell the user exactly this**:',
          '"Connection shell created successfully. Please go to make.powerautomate.com \u2192',
          'Data \u2192 Connections \u2192 find the new connection \u2192 click Authorize."',
          '',
          '## Non-OAuth Connectors (HTTP, SQL with credentials):',
          'These return `connectionStatus: "Connected"` immediately. No manual step needed.',
          '',
          '## Common Connector IDs:',
          '| Connector | connectorId | Needs Manual Auth? |',
          '|-----------|------------|-------------------|',
          '| Office 365 Outlook | shared_office365 | Yes |',
          '| SharePoint | shared_sharepointonline | Yes |',
          '| Teams | shared_teams | Yes |',
          '| OneDrive for Business | shared_onedrive | Yes |',
          '| Excel Online (Business) | shared_excelonlinebusiness | Yes |',
          '| Dataverse | shared_commondataserviceforapps | Yes |',
          '| Approvals | shared_approvals | No |',
          '| HTTP | shared_webcontents | No |',
          '',
          '## PARAMETER FORMAT:',
          'Always use camelCase: connectorId, environmentId, connectionParameters',
          'NOT snake_case: connector_id, environment_id, connection_parameters',
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

// =============================================================================
// Resource Content Generators (for resources/read)
// Replicates src/skills/resources.ts content for REST transport
// =============================================================================
function generateParameterConventions(uri: string): any {
  return {
    contents: [{
      uri,
      mimeType: 'text/markdown',
      text: [
        '# Parameter Naming Conventions \u2014 Power Automate MCP',
        '',
        '## All tool parameters use camelCase:',
        '',
        '| Correct (camelCase) | Wrong (snake_case) | Used By |',
        '|---------------------|-------------------|---------|',
        '| connectorId | connector_id | pa-create-connection, pa-get-connection |',
        '| environmentId | environment_id | Most tools (optional) |',
        '| flowId | flow_id | Flow tools |',
        '| connectionId | connection_id | Connection tools |',
        '| solutionId | solution_id | Solution tools |',
        '| displayName | display_name | pa-create-flow, pa-update-flow |',
        '| connectionReferences | connection_references | pa-create-flow |',
        '| connectionParameters | connection_parameters | pa-create-connection |',
        '| solutionUniqueName | solution_unique_name | pa-export-solution |',
        '| componentId | component_id | pa-add-solution-component |',
        '| componentType | component_type | pa-add-solution-component |',
        '| friendlyName | friendly_name | pa-create-solution |',
        '| uniqueName | unique_name | pa-create-solution |',
        '| publisherId | publisher_id | pa-create-solution |',
        '',
        '## Exception:',
        '- `user_id` uses snake_case (auth tools: pa-auth-start, pa-auth-poll, pa-create-connection)',
        '',
        '## Auto-Resolution:',
        'The server includes a param-resolver that maps snake_case to camelCase automatically.',
        'Agents SHOULD always send camelCase for consistency.',
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
        '',
        'pa-create-connection creates a "shell" connection:',
        '  Status: Error, needsConsent: true',
        '  User authorizes at make.powerautomate.com',
        '  Status changes to: Connected',
        '',
        'IMPORTANT: "Error" + needsConsent is EXPECTED BEHAVIOR, not a failure!',
        '',
        '## Non-OAuth Connectors (HTTP, SQL with connection string):',
        'pa-create-connection returns "Connected" immediately. No manual step needed.',
        '',
        '## Connector Reference:',
        '| Connector | connectorId | Manual Auth? |',
        '|-----------|------------|-------------|',
        '| Office 365 Outlook | shared_office365 | Yes |',
        '| SharePoint | shared_sharepointonline | Yes |',
        '| Teams | shared_teams | Yes |',
        '| OneDrive for Business | shared_onedrive | Yes |',
        '| Dataverse | shared_commondataserviceforapps | Yes |',
        '| Approvals | shared_approvals | No |',
        '| HTTP | shared_webcontents | No |',
        '',
        '## Error Codes:',
        '- 404: Connector name wrong or does not exist',
        '- 403: Admin consent for PowerApps Service needed in Azure AD',
        '- 409: Connection name already exists',
        '- 401: User token expired, re-authenticate with pa-auth-start',
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
        '## Production Environment:',
        `- Environment ID: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not configured)'}`,
        `- Dataverse URL: ${process.env.DATAVERSE_URL || '(not configured)'}`,
        `- Dataverse Scope: https://${process.env.DATAVERSE_URL || 'not-configured'}/.default`,
        `- Tenant: ${process.env.AZURE_TENANT_ID ? 'Configured' : 'Not configured'}`,
        '',
        '## IMPORTANT \u2014 Environment Discovery:',
        '- pa-list-environments may return EMPTY if the service principal lacks Power Platform Admin role',
        `- You do NOT need to call pa-list-environments. The environment ID is: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not set)'}`,
        '- All tools default to this environment automatically when environmentId is omitted',
        '- Just call pa-list-flows, pa-create-flow, etc. directly without environmentId',
        '',
        '## Authentication Architecture:',
        'Service Principal Token (AzureTokenManager):',
        '  - Auto-acquired on boot, auto-refreshed',
        '  - Used for READ ops: list, get, enable/disable',
        '  - 4 scopes: BAP, Flow, PowerApps, Dataverse',
        '',
        'User Delegated Token (UserAuthManager):',
        '  - Device Code Flow per user',
        '  - Used for WRITE ops: create, update, delete',
        '  - Refresh token rotation (~90 day lifetime)',
        '',
        '## Available Capabilities:',
        '- Tools: 26 | Prompts: 3 | Resources: 3',
      ].join('\n'),
    }],
  };
}

const RESOURCE_GENERATORS: Record<string, (uri: string) => any> = {
  'power-automate://docs/parameter-conventions': generateParameterConventions,
  'power-automate://docs/connection-lifecycle': generateConnectionLifecycle,
  'power-automate://config/environment': generateEnvironmentInfo,
};

// =============================================================================
// Main Handler
//
// Processes skill-related JSON-RPC requests for REST transport.
// Returns a JSON-RPC response for skill methods, or null to fall through
// to other handlers (tools/call, initialize, etc.).
// =============================================================================
export function processSkillRequest(method: string, params: any, id: any): any | null {
  switch (method) {
    case 'prompts/list':
      console.log('[REST skills] prompts/list');
      return {
        jsonrpc: '2.0',
        id,
        result: { prompts: PROMPT_REGISTRY },
      };

    case 'prompts/get': {
      const promptName = params?.name;
      console.log(`[REST skills] prompts/get: ${promptName}`);
      const generator = PROMPT_GENERATORS[promptName];
      if (!generator) {
        return {
          jsonrpc: '2.0',
          id,
          error: {
            code: -32602,
            message: `Unknown prompt: ${promptName}. Available: ${Object.keys(PROMPT_GENERATORS).join(', ')}`,
          },
        };
      }
      return {
        jsonrpc: '2.0',
        id,
        result: generator(params?.arguments || {}),
      };
    }

    case 'resources/list':
      console.log('[REST skills] resources/list');
      return {
        jsonrpc: '2.0',
        id,
        result: { resources: RESOURCE_REGISTRY },
      };

    case 'resources/read': {
      const uri = params?.uri;
      console.log(`[REST skills] resources/read: ${uri}`);
      const generator = RESOURCE_GENERATORS[uri];
      if (!generator) {
        return {
          jsonrpc: '2.0',
          id,
          error: {
            code: -32602,
            message: `Unknown resource: ${uri}. Available: ${Object.keys(RESOURCE_GENERATORS).join(', ')}`,
          },
        };
      }
      return {
        jsonrpc: '2.0',
        id,
        result: generator(uri),
      };
    }

    default:
      return null; // Not a skill request — fall through to other handlers
  }
}
