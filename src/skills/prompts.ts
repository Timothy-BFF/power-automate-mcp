/**
 * MCP Skill Prompts — Power Automate MCP (SkillEngine v1.0.0)
 *
 * Workflow guides that agents read BEFORE calling tools.
 * These reduce trial-and-error by providing step-by-step instructions
 * for multi-step operations like authentication, flow creation, and
 * connection creation (including OAuth consent expectations).
 *
 * Registered prompts:
 *   - workflow-auth: Device Code Flow authentication sequence
 *   - workflow-create-flow: Mandatory flow creation procedure + definition format
 *   - workflow-create-connection: Connection creation with OAuth consent guide
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';

export function registerPrompts(server: McpServer): void {

  // =========================================================================
  // Prompt 1: workflow-auth
  // =========================================================================
  server.prompt(
    'workflow-auth',
    'Step-by-step guide for authenticating a user via Device Code Flow. Read this BEFORE calling pa-auth-start.',
    {
      user_id: z.string().optional().describe('User email address (e.g., jose@company.com)')
    },
    async (args) => ({
      messages: [{
        role: 'user' as const,
        content: {
          type: 'text' as const,
          text: [
            '# Authentication Workflow — Device Code Flow',
            '',
            '## When to Use:',
            'Any write operation (create flow, create connection, update flow) requires user authentication.',
            'Read operations (list flows, list environments, get details) use the service principal — no user auth needed.',
            '',
            '## Steps:',
            `1. **Start auth**: Call \`pa-auth-start\` with user_id: "${args.user_id || '<user-email@company.com>'}"`,
            '2. **Show code**: Tell the user exactly:',
            '   "Please go to https://microsoft.com/devicelogin and enter code: {user_code}"',
            '3. **Wait**: Give the user 30–60 seconds to complete sign-in in their browser',
            '4. **Poll**: Call `pa-auth-poll` with the same user_id',
            '   - Status **"pending"** → User hasn\'t completed sign-in yet. Wait 15–30 seconds and poll again.',
            '   - Status **"authenticated"** → Success! Proceed with write operations.',
            '   - Status **"error"** → Auth failed. Start over with pa-auth-start.',
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
            '## IMPORTANT — Environment Discovery:',
            '- `pa-list-environments` may return empty — this is a known limitation (BAP admin role)',
            '- You do NOT need to discover the environment ID — it is configured server-side',
            '- All tools default to the production environment automatically when environmentId is omitted',
            '- Just call pa-list-flows, pa-create-flow, etc. directly without passing environmentId',
            '',
            '## Multi-Scope Token Exchange:',
            'The server automatically exchanges the Flow-scoped token for PowerApps-scoped tokens',
            'when needed (e.g., for pa-create-connection). Users only authenticate once.',
            '',
            '## Parameter Format:',
            '- user_id: Full email address (e.g., "Timothy.Escamilla@bolthousefresh.com")',
          ].join('\n')
        }
      }]
    })
  );

  // =========================================================================
  // Prompt 2: workflow-create-flow (with definition JSON format guide)
  // =========================================================================
  server.prompt(
    'workflow-create-flow',
    'MANDATORY procedure for creating a Power Automate flow. Includes definition JSON format requirements. Agents MUST follow this exact sequence.',
    {
      displayName: z.string().optional().describe('Name of the flow to create'),
      user_id: z.string().optional().describe('Authenticated user email'),
    },
    async (args) => ({
      messages: [{
        role: 'user' as const,
        content: {
          type: 'text' as const,
          text: [
            '# Flow Creation — Mandatory Procedure',
            '',
            `**Flow**: ${args.displayName || '<provide-display-name>'}`,
            `**User**: ${args.user_id || '<must-be-authenticated>'}`,
            '',
            '## Steps (MUST follow in order):',
            '',
            '### Step 1: Authenticate',
            '- Call `pa-auth-start` with user_id → receive device code',
            '- User signs in at microsoft.com/devicelogin',
            '- Call `pa-auth-poll` → wait for status "authenticated"',
            '- See prompt: workflow-auth for full details',
            '',
            '### Step 2: Create Flow (SINGLE CALL — FULL DEFINITION)',
            '- Call `pa-create-flow` with the COMPLETE definition in ONE call',
            '- Required parameters:',
            '  - `displayName`: Human-readable flow name',
            '  - `definition`: Complete workflow definition JSON (see format below)',
            '- Optional parameters:',
            '  - `connectionReferences`: Map of connection references for connectors used',
            '  - `state`: "Started" or "Stopped" (defaults to "Stopped" — safe for testing)',
            '  - `environmentId`: Uses production environment if omitted',
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
            '  ✅ triggerConditions  ❌ trigger_conditions',
            '  ✅ operationId       ❌ operation_id',
            '  ✅ contentType       ❌ content_type',
            '  ✅ retryPolicy       ❌ retry_policy',
            '  ✅ operationOptions  ❌ operation_options',
            '  ✅ splitOn           ❌ split_on',
            '  ✅ trackedProperties ❌ tracked_properties',
            '',
            'Example CORRECT action definition:',
            '  "HTTP_GetData": {',
            '    "type": "Http",',
            '    "runAfter": { "Initialize_Variable": ["Succeeded"] },',
            '    "inputs": {',
            '      "method": "GET",',
            '      "uri": "https://api.example.com/data",',
            '      "headers": { "Content-Type": "application/json" }',
            '    }',
            '  }',
            '',
            'The server auto-normalizes run_after → runAfter and injects missing',
            '$schema/contentVersion as a safety net, but correct definitions avoid errors.',
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
            '## ⚠️ CRITICAL RULES:',
            '',
            '### NEVER Delete on Empty Definition',
            'If `pa-get-flow-details` shows an empty/missing definition but:',
            '- `_fetchedVia` was "admin" → The admin endpoint does NOT return definitions',
            '- `_authType` was "delegated" during creation → The flow WAS created correctly',
            'This is a **known API limitation**, NOT a creation failure. The flow exists and works.',
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
    })
  );

  // =========================================================================
  // Prompt 3: workflow-create-connection
  // =========================================================================
  server.prompt(
    'workflow-create-connection',
    'Complete guide for creating a connector connection, including expected OAuth consent flow. Read BEFORE calling pa-create-connection.',
    {
      connectorId: z.string().optional().describe('Connector name (e.g., shared_office365)'),
      user_id: z.string().optional().describe('Authenticated user email'),
    },
    async (args) => ({
      messages: [{
        role: 'user' as const,
        content: {
          type: 'text' as const,
          text: [
            '# Connection Creation Workflow',
            '',
            `**Connector**: ${args.connectorId || '<provide-connector-id>'}`,
            `**User**: ${args.user_id || '<must-be-authenticated>'}`,
            '',
            '## Pre-requisite:',
            'User MUST be authenticated first (pa-auth-start → pa-auth-poll → "authenticated").',
            'See prompt: workflow-auth for details.',
            '',
            '## Steps:',
            '1. **Create connection**: Call `pa-create-connection` with:',
            `   - \`connectorId\`: "${args.connectorId || 'shared_office365'}" (exact connector short name)`,
            `   - \`user_id\`: "${args.user_id || '<user-email@company.com>'}"`,
            '   - `environmentId`: Optional — uses production environment if omitted',
            '   - `connectionParameters`: Optional — only needed for SQL, HTTP, custom connectors',
            '',
            '## ⚠️ EXPECTED BEHAVIOR — OAuth Connectors:',
            '',
            'For **Office 365, SharePoint, Teams, Outlook, OneDrive**:',
            '- The API response WILL show: `connectionStatus: "Error"`, `needsConsent: true`',
            '- **This is 100% NORMAL — not a bug, not a failure!**',
            '- The API creates a "shell" connection requiring interactive browser authorization',
            '',
            '**Tell the user exactly this**:',
            '"Connection shell created successfully. Please go to make.powerautomate.com →',
            'Data → Connections → find the new connection → click Authorize."',
            '',
            'After the user authorizes in their browser, the status becomes "Connected".',
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
            '',
            'The server includes a param-resolver that maps snake_case → camelCase automatically,',
            'but agents SHOULD always send camelCase to avoid unnecessary resolution.',
          ].join('\n')
        }
      }]
    })
  );

  console.log('[Skills] 3 workflow prompts registered: workflow-auth, workflow-create-flow, workflow-create-connection');
}
