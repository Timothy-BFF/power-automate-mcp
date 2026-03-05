# Power Automate MCP Server

Model Context Protocol (MCP) server for Microsoft Power Automate administration.

**Built for:** Bolthouse Fresh Foods
**Architected by:** MCA (Model Context Architect)
**Deployed on:** Railway

## Tools

| Tool | Description |
|------|-------------|
| `pa-list-environments` | List all Power Platform environments |
| `pa-list-flows` | List flows in an environment |
| `pa-get-flow-details` | Get detailed flow information |
| `pa-get-run-history` | Get flow run history |
| `pa-get-run-details` | Get specific run details |
| `pa-list-connections` | List environment connections |
| `pa-enable-disable-flow` | Start or stop a flow |
| `pa-delete-flow` | Permanently delete a flow |
| `pa-trigger-flow` | Manually trigger a flow |
| `pa-cancel-run` | Cancel a running flow |

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | Yes | Service principal app ID |
| `AZURE_CLIENT_SECRET` | Yes | Service principal secret |
| `SIMTHEORY_AUTH_TOKEN` | Yes | Simtheory.ai auth token |
| `DEFAULT_ENVIRONMENT_ID` | No | Default Power Platform environment |
| `PORT` | No | Server port (default: 8080) |
| `LOG_LEVEL` | No | Logging level (default: info) |

## Authentication

Uses OAuth 2.0 client credentials flow with automatic token refresh.
The service principal must be registered via `New-PowerAppManagementApp` in Power Platform.
