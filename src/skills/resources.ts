/**
 * MCP Skill Resources — Power Automate MCP (SkillEngine v1.0.0)
 *
 * Read-only reference documents that agents access for context.
 * These provide parameter conventions, lifecycle guides, and environment info
 * so agents can make informed decisions without trial-and-error.
 *
 * Registered resources:
 *   - parameter-conventions (power-automate://docs/parameter-conventions)
 *   - connection-lifecycle  (power-automate://docs/connection-lifecycle)
 *   - environment-info      (power-automate://config/environment) [dynamic]
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerResources(server: McpServer): void {

  // =========================================================================
  // Resource 1: parameter-conventions
  // Prevents: snake_case/camelCase parameter mismatches across all tools.
  // Root cause of Jose's 5-round debugging saga with pa-create-connection.
  // =========================================================================
  server.resource(
    'parameter-conventions',
    'power-automate://docs/parameter-conventions',
    async (uri) => ({
      contents: [{
        uri: uri.href,
        mimeType: 'text/markdown',
        text: [
          '# Parameter Naming Conventions — Power Automate MCP',
          '',
          '## All tool parameters use camelCase:',
          '',
          '| ✅ Correct (camelCase) | ❌ Wrong (snake_case) | Used By |',
          '|----------------------|---------------------|---------|',
          '| connectorId | connector_id | pa-create-connection, pa-get-connection |',
          '| environmentId | environment_id | Most tools (optional) |',
          '| flowId | flow_id | Flow tools (get, update, enable, disable, delete, trigger) |',
          '| connectionId | connection_id | pa-get-connection, pa-delete-connection |',
          '| solutionId | solution_id | Solution tools (get, delete, list-components) |',
          '| displayName | display_name | pa-create-flow, pa-update-flow |',
          '| connectionReferences | connection_references | pa-create-flow, pa-update-flow |',
          '| connectionParameters | connection_parameters | pa-create-connection |',
          '| solutionUniqueName | solution_unique_name | pa-export-solution, pa-add-solution-component |',
          '| componentId | component_id | pa-add-solution-component |',
          '| componentType | component_type | pa-add-solution-component |',
          '| addRequiredComponents | add_required_components | pa-add-solution-component |',
          '| includeManaged | include_managed | pa-list-solutions |',
          '| friendlyName | friendly_name | pa-create-solution |',
          '| uniqueName | unique_name | pa-create-solution |',
          '| publisherId | publisher_id | pa-create-solution |',
          '',
          '## Exception:',
          '- `user_id` uses snake_case (auth tools: pa-auth-start, pa-auth-poll, pa-create-connection)',
          '',
          '## Value Formats:',
          '- **user_id**: Full email address — "jose@company.com"',
          '- **flowId**: GUID — "a1b2c3d4-e5f6-7890-abcd-ef1234567890"',
          '- **connectorId**: Short name — "shared_office365", "shared_sql", "shared_sharepointonline"',
          '  Or full API path — "/providers/Microsoft.PowerApps/apis/shared_office365"',
          '- **solutionId**: GUID from pa-list-solutions or pa-get-solution',
          '- **environmentId**: GUID — "c8113fbe-e00c-ef03-9ba5-548acf1f5807"',
          '- **connectionId**: Name field from pa-list-connections (GUID format)',
          '',
          '## Auto-Resolution:',
          'The server includes a param-resolver that maps snake_case → camelCase automatically.',
          'This means `connector_id` will work, but agents SHOULD always send camelCase',
          'for consistency and to match the tool schema definitions.',
        ].join('\n'),
      }],
    })
  );

  // =========================================================================
  // Resource 2: connection-lifecycle
  // Prevents: agents interpreting OAuth "Error" status as a failure.
  // =========================================================================
  server.resource(
    'connection-lifecycle',
    'power-automate://docs/connection-lifecycle',
    async (uri) => ({
      contents: [{
        uri: uri.href,
        mimeType: 'text/markdown',
        text: [
          '# Connection Lifecycle Guide',
          '',
          '## OAuth Connectors (Office 365, SharePoint, Teams, Outlook, OneDrive):',
          '',
          '```',
          'pa-create-connection → "shell" created → Status: Error, needsConsent: true',
          '                                          ↓',
          '                     User authorizes at make.powerautomate.com',
          '                                          ↓',
          '                     Status changes to: Connected ✅',
          '```',
          '',
          '1. `pa-create-connection` → API creates a "shell" connection',
          '2. Response: connectionStatus = "Error", needsConsent = true',
          '3. ⚠️ **This is EXPECTED BEHAVIOR — not a failure!**',
          '4. User must manually authorize at: make.powerautomate.com → Data → Connections → Authorize',
          '5. After browser authorization: status becomes "Connected"',
          '',
          '## Non-OAuth Connectors (HTTP, SQL with connection string):',
          '',
          '```',
          'pa-create-connection → Immediately "Connected" ✅',
          '```',
          '',
          '1. `pa-create-connection` with connectionParameters (server, database, etc.)',
          '2. Response: connectionStatus = "Connected"',
          '3. No manual step required',
          '',
          '## Connector Reference Table:',
          '',
          '| Connector | connectorId | Type | Manual Auth? |',
          '|-----------|------------|------|-------------|',
          '| Office 365 Outlook | shared_office365 | OAuth | ✅ Yes |',
          '| SharePoint | shared_sharepointonline | OAuth | ✅ Yes |',
          '| Teams | shared_teams | OAuth | ✅ Yes |',
          '| OneDrive for Business | shared_onedrive | OAuth | ✅ Yes |',
          '| Excel Online (Business) | shared_excelonlinebusiness | OAuth | ✅ Yes |',
          '| Dataverse | shared_commondataserviceforapps | OAuth | ✅ Yes |',
          '| Dynamics 365 | shared_dynamicscrmonline | OAuth | ✅ Yes |',
          '| Approvals | shared_approvals | Service | ❌ No |',
          '| HTTP | shared_webcontents | API Key | ❌ No |',
          '| Notifications | shared_flowpush | Service | ❌ No |',
          '| RSS | shared_rss | None | ❌ No |',
          '',
          '## Connection Ownership:',
          '- Connections are owned by the authenticated user (via pa-auth-start)',
          '- Each user creates their own connections — connections are NOT shared by default',
          '- The connection is tied to the user\'s identity in the target environment',
          '',
          '## Error Codes:',
          '- **404**: Connector name is wrong or doesn\'t exist in the environment',
          '- **403**: Admin consent for PowerApps Service may be needed in Azure AD',
          '- **409**: A connection with that exact name already exists (retry with new name)',
          '- **401**: User token expired — re-authenticate with pa-auth-start',
        ].join('\n'),
      }],
    })
  );

  // =========================================================================
  // Resource 3: environment-info (DYNAMIC — reads env vars at request time)
  // Provides current environment context so agents know where they're operating.
  // =========================================================================
  server.resource(
    'environment-info',
    'power-automate://config/environment',
    async (uri) => ({
      contents: [{
        uri: uri.href,
        mimeType: 'text/markdown',
        text: [
          '# Current Environment Configuration',
          '',
          '## Production Environment:',
          `- **Environment ID**: ${process.env.POWER_PLATFORM_ENVIRONMENT_ID || '(not configured)'}`,
          `- **Dataverse URL**: ${process.env.DATAVERSE_URL || '(not configured)'}`,
          `- **Dataverse Scope**: https://${process.env.DATAVERSE_URL || 'not-configured'}/.default`,
          `- **Tenant**: ${process.env.AZURE_TENANT_ID ? 'Configured (' + process.env.AZURE_TENANT_ID.substring(0, 8) + '...)' : 'Not configured'}`,
          '',
          '## Authentication Architecture:',
          '```',
          '┌─────────────────────────────┐   ┌──────────────────────────────┐',
          '│   Service Principal Token   │   │   User Delegated Token       │',
          '│   (AzureTokenManager)       │   │   (UserAuthManager)          │',
          '│                             │   │                              │',
          '│   • Auto-acquired on boot   │   │   • Device Code Flow         │',
          '│   • Auto-refreshed          │   │   • Per-user token store     │',
          '│   • Used for READ ops       │   │   • Multi-scope exchange     │',
          '│                             │   │   • Used for WRITE ops       │',
          '└─────────────┬───────────────┘   └──────────────┬───────────────┘',
          '              │                                   │',
          '    list, get, enable/disable            create, update, delete',
          '    run history, solutions               connections, flows',
          '```',
          '',
          '## Admin Token Scopes (Service Principal):',
          '| Scope | Purpose |',
          '|-------|---------|',
          '| api.bap.microsoft.com/.default | Environment listing (BAP admin) |',
          '| service.flow.microsoft.com/.default | Flow management (list, get, enable/disable) |',
          '| service.powerapps.com/.default | Connection listing (admin) |',
          `| ${process.env.DATAVERSE_URL || 'dataverse'}/.default | Solution operations (CRUD) |`,
          '',
          '## User Token Scopes (Delegated via Device Code Flow):',
          '| Scope | Purpose | Acquired Via |',
          '|-------|---------|-------------|',
          '| service.flow.microsoft.com/.default | Flow creation/update | Device Code Flow (primary) |',
          '| service.powerapps.com/.default | Connection creation | Refresh token exchange |',
          '',
          '## Available Capabilities:',
          '- **Tools**: 26 (environments, flows, runs, connections, solutions, auth)',
          '- **Prompts**: 3 (workflow-auth, workflow-create-flow, workflow-create-connection)',
          '- **Resources**: 3 (parameter-conventions, connection-lifecycle, this document)',
        ].join('\n'),
      }],
    })
  );

  console.log('[Skills] 3 knowledge resources registered: parameter-conventions, connection-lifecycle, environment-info');
}
