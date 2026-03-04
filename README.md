# Power Automate MCP Server

> Model Context Protocol server for managing Microsoft Power Automate flows via the Power Platform APIs.
> Designed for [Simtheory.ai](https://simtheory.ai) integration with Railway deployment.

**Author:** GROW by Bolthouse Fresh (Architected by MCA)

---

## Architecture Overview

```
Simtheory.ai Workspace
  └── Power Automate MCP Server (Railway)
        ├── Azure AD Token Manager ("No-Bother" Protocol)
        │     ├── Flow API scope (service.flow.microsoft.com)
        │     └── Management API scope (management.azure.com)
        ├── Flow Client       → api.flow.microsoft.com
        ├── Environment Client → api.bap.microsoft.com
        └── Connection Client  → api.flow.microsoft.com
```

## Tools (10 Production-Ready)

| Tool | Description |
|------|-------------|
| `pa-list-flows` | List all flows in an environment |
| `pa-get-flow-details` | Get complete flow details including definition |
| `pa-enable-disable-flow` | Start or stop a flow |
| `pa-delete-flow` | Permanently delete a flow (with confirmation) |
| `pa-trigger-flow` | Trigger a flow via HTTP request trigger |
| `pa-get-run-history` | Get execution history for a flow |
| `pa-get-run-details` | Get details of a specific run |
| `pa-cancel-run` | Cancel a running flow execution |
| `pa-list-environments` | List all Power Platform environments |
| `pa-list-connections` | List all API connections in an environment |

## Prerequisites

### 1. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click **New registration**
3. Name: `Power Automate MCP Server`
4. Supported account types: **Single tenant**
5. Click **Register**
6. Note your **Application (client) ID** and **Directory (tenant) ID**
7. Go to **Certificates & secrets** → New client secret → Note the **Value**
8. Go to **API permissions** → Add permissions:
   - **Flow Service** (`https://service.flow.microsoft.com/`)
     - `Flows.Read.All`
     - `Flows.Manage.All`
   - **Power Platform** (via Microsoft Graph or direct)
   - Click **Grant admin consent**

### 2. Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `PORT` | No | Server port (default: 8080) |
| `LOG_LEVEL` | No | Logging level (default: info) |
| `SIMTHEORY_AUTH_TOKEN` | **Yes** | Simtheory.ai authorization token |
| `AZURE_TENANT_ID` | **Yes** | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | **Yes** | Azure AD application client ID |
| `AZURE_CLIENT_SECRET` | **Yes** | Azure AD client secret |
| `POWER_PLATFORM_ENVIRONMENT_ID` | No | Default Power Platform environment ID |

## Deployment (Railway)

### Step 1: Connect Repository

1. Log into [Railway](https://railway.app)
2. **New Project** → **Deploy from GitHub Repo**
3. Select `Timothy-BFF/power-automate-mcp`

### Step 2: Set Environment Variables

In Railway project settings → Variables, add all required environment variables listed above.

### Step 3: Deploy

Railway will auto-detect the Dockerfile and deploy. Verify:

```bash
curl https://your-app.up.railway.app/health
```

Expected response:
```json
{
  "status": "healthy",
  "server": "power-automate-mcp",
  "version": "1.0.0",
  "auth": {
    "flowToken": "ok",
    "managementToken": "ok"
  },
  "tools": 10
}
```

## Simtheory.ai Integration

### MCP Server Configuration

| Setting | Value |
|---------|-------|
| **Server URL** | `https://your-railway-app.up.railway.app/sse` |
| **Transport** | `sse` |
| **Authorization Header** | `Bearer <your-simtheory-auth-token>` |

## Example Workflows

### List all flows, then check a specific flow's run history
```
"List all my Power Automate flows"
→ pa-list-flows

"Show me the last 5 runs for the Invoice Processing flow"
→ pa-get-run-history (flowId, top: 5)

"Why did the last run fail?"
→ pa-get-run-details (flowId, runId)
```

### Trigger a flow with input data
```
"Trigger the Employee Onboarding flow for John Smith in Engineering"
→ pa-trigger-flow (flowId, body: { name: "John Smith", department: "Engineering" })
```

### Manage flow state
```
"Disable the Daily Report flow — we're doing maintenance this weekend"
→ pa-enable-disable-flow (flowId, action: "disable")

"Re-enable it now"
→ pa-enable-disable-flow (flowId, action: "enable")
```

## Token Management (No-Bother Protocol)

The server implements automatic Azure AD token lifecycle management:

- **Dual-scope support:** Separate tokens for Flow API and Management API
- **Proactive refresh:** Tokens refreshed when < 5 minutes remain
- **401 retry:** Automatic refresh + retry on authentication failures
- **429 backoff:** Respects rate limit headers from Power Platform APIs
- **Thread-safe:** Concurrent requests share refresh promises

Users never need to manually refresh tokens.

## Project Structure

```
src/
├── index.ts                    # Main server entry point
├── config/
│   └── index.ts                # Environment variable configuration
├── auth/
│   └── azure-token-manager.ts  # No-Bother Token Protocol
├── clients/
│   ├── flow-client.ts          # Flow Management API wrapper
│   ├── environment-client.ts   # Environment API wrapper
│   └── connection-client.ts    # Connection API wrapper
├── mcp/
│   └── tools/
│       ├── list-flows.ts
│       ├── get-flow-details.ts
│       ├── enable-disable-flow.ts
│       ├── delete-flow.ts
│       ├── trigger-flow.ts
│       ├── get-run-history.ts
│       ├── get-run-details.ts
│       ├── cancel-run.ts
│       ├── list-environments.ts
│       └── list-connections.ts
└── utils/
    └── logger.ts               # Winston logging
```

## License

MIT — GROW by Bolthouse Fresh
