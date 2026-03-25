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
│  │  • Auto-refresh      │   │  • Per-user tokens         │  │
│  │  • 4 scopes          │   │  • Refresh token rotation  │  │
│  └──────────┬───────────┘   └──────────────┬─────────────┘  │
│             │                               │               │
│       READ OPERATIONS                 WRITE OPERATIONS      │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  SkillEngine v1.0.0                                  │   │
│  │  3 Workflow Prompts + 3 Knowledge Resources          │   │
│  │  Available via SSE native + REST JSON-RPC             │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Definition Normalizer                                │   │
│  │  Auto-fixes: run_after→runAfter, injects $schema     │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                             │
│  Transports: SSE (/sse) + REST JSON-RPC (/mcp)             │
└─────────────────────────────────────────────────────────────┘
```

---

## Tools (26)

### Environments
| Tool | Description |
|------|-------------|
| `pa-list-environments` | List Power Platform environments (may return empty — see Known Limitations) |

### Flows
| Tool | Description |
|------|-------------|
| `pa-list-flows` | List flows in an environment |
| `pa-get-flow-details` | Get full flow details including definition |
| `pa-create-flow` | Create a cloud flow (auto-normalizes definition) |
| `pa-update-flow` | Update flow definition, name, or state (auto-normalizes) |
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
| `pa-list-connections` | List all connections (admin) |
| `pa-get-connection` | Get connection details |
| `pa-delete-connection` | Delete a connection |
| `pa-create-connection` | Create a connector connection (requires user auth) |

### Solutions (Dataverse)
| Tool | Description |
|------|-------------|
| `pa-list-solutions` | List solutions (managed/unmanaged) |
| `pa-get-solution` | Get solution details including publisher |
| `pa-create-solution` | Create a new unmanaged solution |
| `pa-delete-solution` | Delete a solution |
| `pa-list-solution-components` | List components in a solution |
| `pa-export-solution` | Export solution as ZIP (base64) |
| `pa-add-solution-component` | Add a component to a solution |

### Authentication
| Tool | Description |
|------|-------------|
| `pa-auth-start` | Start Device Code Flow for a user |
| `pa-auth-poll` | Poll for authentication completion |
| `pa-auth-status` | Check current auth state |

---

## Skills — SkillEngine v1.0.0

Skills are **embedded operational knowledge** that agents read BEFORE calling tools. Available via both SSE and REST transports.

### Workflow Prompts (3)

| Prompt | Purpose | Prevents |
|--------|---------|----------|
| `workflow-auth` | Device Code Flow sequence + environment discovery note | Auth errors, empty environment confusion |
| `workflow-create-flow` | Mandatory 4-step procedure + **definition JSON format** | Empty definitions, `run_after` errors, missing `$schema` |
| `workflow-create-connection` | Connection creation with OAuth consent guide | Agents panicking at expected "Error" status |

### Knowledge Resources (3)

| Resource | URI | Purpose |
|----------|-----|----------|
| `parameter-conventions` | `power-automate://docs/parameter-conventions` | camelCase mapping for tool params + definition body |
| `connection-lifecycle` | `power-automate://docs/connection-lifecycle` | OAuth vs non-OAuth connector behavior |
| `environment-info` | `power-automate://config/environment` | Current env ID, auth architecture (dynamic) |

---

## Flow Definition Normalizer

The server automatically fixes common AI-generated definition errors:

| Fix | Example | Applied To |
|-----|---------|------------|
| `run_after` → `runAfter` | Recursive in all actions | `pa-create-flow`, `pa-update-flow` |
| Missing `$schema` | Injects Logic Apps schema URL | `pa-create-flow`, `pa-update-flow` |
| Missing `contentVersion` | Injects `1.0.0.0` | `pa-create-flow`, `pa-update-flow` |
| `trigger_conditions` → `triggerConditions` | Recursive | `pa-create-flow`, `pa-update-flow` |
| `operation_id` → `operationId` | Recursive | `pa-create-flow`, `pa-update-flow` |
| `retry_policy` → `retryPolicy` | Recursive | `pa-create-flow`, `pa-update-flow` |

Diagnostics logged to Railway: `[NormalizeDef] Applied 3 fixes: injected $schema, injected contentVersion, remapped 12 snake_case properties`

---

## Authentication

### Dual-Token Architecture

**Service Principal** (automatic):
- Acquired on startup via client credentials
- Used for READ operations (list, get, enable/disable)
- 4 scopes: BAP, Flow, PowerApps, Dataverse

**User Delegated** (per-user Device Code Flow):
- Initiated via `pa-auth-start` → user signs in at `microsoft.com/devicelogin`
- Used for WRITE operations (create, update, delete)
- Refresh token rotation (~90 day lifetime)
- Multi-scope exchange: Flow token exchanged for PowerApps token when needed

---

## Known Limitations

| Limitation | Impact | Workaround |
|-----------|--------|------------|
| `pa-list-environments` returns empty | Service principal may lack Power Platform Admin role | All tools default to configured environment ID — skip this call |
| OAuth connections show `Error` status | Expected behavior for Office 365, SharePoint, etc. | User authorizes at `make.powerautomate.com` → Data → Connections |
| Admin endpoint doesn't return definitions | `pa-get-flow-details` may show empty definition | This is an API limitation, not a creation failure |

---

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | ✅ | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | ✅ | App registration client ID |
| `AZURE_CLIENT_SECRET` | ✅ | App registration client secret |
| `POWER_PLATFORM_ENVIRONMENT_ID` | ✅ | Target Power Platform environment ID |
| `DATAVERSE_URL` | ✅ | Dataverse instance URL |
| `PORT` | ❌ | Server port (default: 8080) |

---

## Deployment (Railway)

Push to `main` → Railway auto-builds → auto-deploys.

### Endpoints

| Endpoint | Protocol | Purpose |
|----------|----------|----------|
| `GET /sse` | SSE | MCP Server-Sent Events transport |
| `POST /messages` | SSE | MCP message handler |
| `POST /mcp` | REST | JSON-RPC 2.0 (tools + skills) |
| `POST /` | REST | JSON-RPC 2.0 (alias) |
| `GET /health` | HTTP | Health check |

---

## Project Structure

```
src/
├── index.ts                        # Main entry: Express + MCP + tools + skills
├── auth/
│   ├── azure-token-manager.ts      # Service principal token lifecycle
│   └── user-auth-manager.ts        # Per-user Device Code Flow
├── api/
│   └── power-platform-client.ts     # HTTP client for Flow + PowerApps APIs
├── clients/
│   └── solution-client.ts          # Dataverse Web API client
├── config/
│   └── environment-resolver.ts     # Environment ID resolution
├── tools/
│   └── tool-descriptions.ts        # Centralized tool descriptions
├── utils/
│   ├── param-resolver.ts           # snake_case → camelCase param unwrapper
│   └── normalize-definition.ts     # Flow definition auto-fixer
└── skills/
    ├── register.ts                 # SkillEngine entry point
    ├── prompts.ts                  # 3 workflow prompts (SSE)
    ├── resources.ts                # 3 knowledge resources (SSE)
    └── rest-skills.ts              # REST JSON-RPC skill handler
```

---

## Azure AD App Registration

### Delegated Permissions
| API | Permission |
|-----|------------|
| Microsoft Graph | `User.Read` |
| Flow Service | `Flows.Manage.All` |
| PowerApps Service | `User` |

### Application Permissions
| API | Permission |
|-----|------------|
| Flow Service | `Flows.Read.All` |
| PowerApps Service | `Connectors.Read.All` |

### Platform Configuration
- Enable `https://login.microsoftonline.com/common/oauth2/nativeclient` (Device Code Flow)
- **Allow public client flows**: Yes

---

## Version History

| Version | Date | Changes |
|---------|------|----------|
| v3.3.0 | 2026-03-25 | SkillEngine v1.0.0, REST skill endpoints, definition normalizer, production env switch, param-resolver, pa-create-solution, pa-delete-solution |
| v3.2.0 | 2026-03-22 | 5 Dataverse solution tools |
| v3.1.0 | 2026-03-21 | pa-get-connection, pa-delete-connection |
| v3.0.0 | 2026-03-20 | Dual-token architecture, Device Code Flow, unified SSE + REST |
| v2.4.0 | 2026-03-17 | Service principal auth, SSE transport |

---

## License

Internal — Bolthouse Fresh Foods
