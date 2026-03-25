# Power Automate MCP v3.3.0

A production-grade **Model Context Protocol (MCP) server** that gives AI agents full control over Microsoft Power Automate — flows, connections, solutions, and environments — through 26 tools and 6 embedded skills.

Deployed on **Railway**. Built with **TypeScript**. Authenticated via **Azure AD** dual-token architecture.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                  Power Automate MCP v3.3.0                   │
│                                                             │
│  ┌──────────────────────┐   ┌────────────────────────────┐  │
│  │  AzureTokenManager   │   │  UserAuthManager            │  │
│  │  (Service Principal)  │   │  (Device Code Flow)        │  │
│  │                      │   │                            │  │
│  │  • Client credentials│   │  • Per-user tokens         │  │
│  │  • Auto-refresh      │   │  • Refresh token rotation  │  │
│  │  • 4 scopes          │   │  • Multi-scope exchange    │  │
│  └──────────┬───────────┘   └──────────────┬─────────────┘  │
│             │                               │               │
│       READ OPERATIONS                 WRITE OPERATIONS      │
│     (admin endpoints)               (user endpoints)        │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  SkillEngine v1.0.0                                  │   │
│  │  3 Workflow Prompts + 3 Knowledge Resources          │   │
│  │  Agents read these BEFORE calling tools               │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                             │
│  Transports: SSE (/sse) + REST JSON-RPC (/mcp)             │
│  Health: /health                                            │
└─────────────────────────────────────────────────────────────┘
```

---

## Tools (26)

### Environments
| Tool | Description |
|------|-------------|
| `pa-list-environments` | List all Power Platform environments in the tenant |

### Flows
| Tool | Description |
|------|-------------|
| `pa-list-flows` | List flows in an environment (filterable) |
| `pa-get-flow-details` | Get full flow details including definition |
| `pa-create-flow` | Create a new cloud flow (requires user auth) |
| `pa-update-flow` | Update flow definition, name, or state |
| `pa-enable-flow` | Enable (start) a flow |
| `pa-disable-flow` | Disable (stop) a flow |
| `pa-delete-flow` | Permanently delete a flow |
| `pa-trigger-flow` | Manually trigger a flow |

### Runs
| Tool | Description |
|------|-------------|
| `pa-get-run-history` | Get recent run history for a flow |
| `pa-get-run-details` | Get details of a specific run |
| `pa-cancel-run` | Cancel a running flow execution |

### Connections
| Tool | Description |
|------|-------------|
| `pa-list-connections` | List all connections in the environment (admin) |
| `pa-get-connection` | Get connection details |
| `pa-delete-connection` | Delete a connection |
| `pa-create-connection` | Create a new connector connection (requires user auth) |

### Solutions (Dataverse)
| Tool | Description |
|------|-------------|
| `pa-list-solutions` | List solutions (managed/unmanaged) |
| `pa-get-solution` | Get solution details including publisher |
| `pa-create-solution` | Create a new unmanaged solution |
| `pa-delete-solution` | Delete a solution |
| `pa-list-solution-components` | List components in a solution |
| `pa-export-solution` | Export solution as ZIP (base64) |
| `pa-add-solution-component` | Add a flow or component to a solution |

### Authentication
| Tool | Description |
|------|-------------|
| `pa-auth-start` | Start Device Code Flow for a user |
| `pa-auth-poll` | Poll for authentication completion |
| `pa-auth-status` | Check current auth state for a user |

---

## Skills — SkillEngine v1.0.0

Skills are **embedded operational knowledge** that agents read BEFORE calling tools. They reduce trial-and-error by providing workflow guides, parameter conventions, and lifecycle documentation.

### Workflow Prompts (3)

Agents access these via `prompts/get` to receive step-by-step instructions.

| Prompt | Purpose | Prevents |
|--------|---------|----------|
| `workflow-auth` | Device Code Flow authentication sequence | Agents calling write tools without auth |
| `workflow-create-flow` | Mandatory 4-step flow creation procedure | Empty definitions, premature deletion, missing verification |
| `workflow-create-connection` | Connection creation with OAuth consent guide | Agents panicking at expected "Error" status for OAuth connectors |

### Knowledge Resources (3)

Agents access these via `resources/read` for reference data.

| Resource | URI | Purpose |
|----------|-----|----------|
| `parameter-conventions` | `power-automate://docs/parameter-conventions` | camelCase vs snake_case mapping for all 26 tools |
| `connection-lifecycle` | `power-automate://docs/connection-lifecycle` | OAuth vs non-OAuth connector behavior, error codes |
| `environment-info` | `power-automate://config/environment` | Current environment ID, Dataverse URL, auth architecture (dynamic) |

---

## Authentication

### Dual-Token Architecture

**Service Principal** (automatic, no user interaction):
- Acquired on startup via client credentials
- Used for all READ operations (list, get, enable/disable)
- 4 scopes: BAP, Flow, PowerApps, Dataverse
- Auto-refreshed when TTL < 5 minutes

**User Delegated** (per-user, Device Code Flow):
- Initiated via `pa-auth-start` → user signs in at `microsoft.com/devicelogin`
- Used for all WRITE operations (create, update, delete flows/connections)
- Refresh token rotation (~90 day lifetime)
- Multi-scope exchange: Flow token silently exchanged for PowerApps token when needed

### Authentication Flow

```
User: "Create a flow"
  → Agent reads prompt: workflow-auth
  → Agent calls: pa-auth-start(user_id: "user@company.com")
  → Server returns: { user_code: "ABCD1234", verification_uri: "microsoft.com/devicelogin" }
  → User enters code in browser, signs in with M365 account
  → Agent calls: pa-auth-poll(user_id: "user@company.com")
  → Server returns: { status: "authenticated" }
  → Agent proceeds with pa-create-flow using the user's delegated token
  → Flow is owned by the user's identity
```

---

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | ✅ | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | ✅ | App registration client ID |
| `AZURE_CLIENT_SECRET` | ✅ | App registration client secret |
| `POWER_PLATFORM_ENVIRONMENT_ID` | ✅ | Target Power Platform environment ID |
| `DATAVERSE_URL` | ✅ | Dataverse instance URL (e.g., `bolthousefreshprod.crm.dynamics.com`) |
| `PORT` | ❌ | Server port (default: 8080, Railway sets automatically) |

---

## Deployment (Railway)

The server deploys automatically from this repo via Railway:

1. Push to `main` → Railway detects → builds TypeScript → deploys
2. Dockerfile handles build: `npm run build` → `node dist/index.js`
3. Health check: `GET /health` returns tools, skills, auth status, Dataverse config

### Endpoints

| Endpoint | Protocol | Purpose |
|----------|----------|----------|
| `GET /sse` | SSE | MCP Server-Sent Events transport |
| `POST /messages?sessionId=...` | SSE | MCP message handler |
| `POST /mcp` | REST | JSON-RPC 2.0 endpoint |
| `POST /` | REST | JSON-RPC 2.0 (alias) |
| `POST /api` | REST | JSON-RPC 2.0 (alias) |
| `POST /tools` | REST | JSON-RPC 2.0 (alias) |
| `GET /health` | HTTP | Health check |

---

## Project Structure

```
src/
├── index.ts                    # Main entry — Express + MCP server + tool registration
├── types.ts                    # TypeScript interfaces
├── build-version.ts            # Version constants
├── auth/
│   ├── azure-token-manager.ts  # Service principal token lifecycle (4 scopes)
│   └── user-auth-manager.ts    # Per-user Device Code Flow + refresh rotation
├── api/
│   └── power-platform-client.ts # HTTP client for Flow + PowerApps APIs
├── clients/
│   └── solution-client.ts      # Dataverse Web API client (solutions CRUD)
├── config/
│   └── environment-resolver.ts # Environment ID resolution from env vars
├── tools/
│   └── tool-descriptions.ts    # Centralized tool descriptions
├── utils/
│   └── param-resolver.ts       # Universal snake_case → camelCase parameter unwrapper
├── skills/
│   ├── register.ts             # SkillEngine entry point — registers prompts + resources
│   ├── prompts.ts              # 3 workflow prompts (auth, create-flow, create-connection)
│   └── resources.ts            # 3 knowledge resources (params, lifecycle, env-info)
└── mcp/
    ├── server.ts               # Legacy v2.0.0 MCP server (superseded by index.ts)
    └── tools/                  # Legacy v2.0.0 tool files (superseded by index.ts)
```

---

## Azure AD App Registration

### Required API Permissions (Delegated)

| API | Permission | Purpose |
|-----|-----------|----------|
| Microsoft Graph | `User.Read` | Basic profile |
| Flow Service | `Flows.Manage.All` | Flow CRUD |
| PowerApps Service | `User` | Connection management |

### Required API Permissions (Application)

| API | Permission | Purpose |
|-----|-----------|----------|
| Flow Service | `Flows.Read.All` | Admin flow listing |
| PowerApps Service | `Connectors.Read.All` | Admin connection listing |

### Platform Configuration

- **Mobile and desktop applications**: Enable `https://login.microsoftonline.com/common/oauth2/nativeclient` (for Device Code Flow)
- **Allow public client flows**: Yes

---

## Version History

| Version | Date | Changes |
|---------|------|----------|
| v3.3.0 | 2026-03-24 | SkillEngine v1.0.0 (3 prompts + 3 resources), production environment switch, pa-create-solution, pa-delete-solution, universal param-resolver, pa-list-connections fix |
| v3.2.0 | 2026-03-22 | 5 Dataverse solution tools |
| v3.1.0 | 2026-03-21 | pa-get-connection, pa-delete-connection |
| v3.0.0 | 2026-03-20 | Dual-token architecture, Device Code Flow, unified SSE + REST |
| v2.4.0 | 2026-03-17 | Service principal auth, SSE transport |

---

## License

Internal — Bolthouse Fresh Foods
