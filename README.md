# Power Automate MCP v3.4.0

A production-grade **Model Context Protocol (MCP) server** that gives AI agents full control over Microsoft Power Automate — flows, connections, solutions, and environments — through 27 tools and 6 embedded skills.

Deployed on **Railway**. Built with **TypeScript**. Authenticated via **Azure AD** dual-token architecture.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                  Power Automate MCP v3.4.0                      │
│                                                                 │
│  ┌──────────────────────┐   ┌────────────────────────────────┐  │
│  │  AzureTokenManager   │   │  UserAuthManager               │  │
│  │  (Service Principal)  │   │  (Device Code Flow)           │  │
│  │  • Auto-refresh      │   │  • Per-user tokens            │  │
│  │  • 4 scopes          │   │  • Refresh token rotation     │  │
│  └──────────┬───────────┘   └──────────────┬─────────────────┘  │
│             │                               │                   │
│       READ OPERATIONS                 WRITE OPERATIONS          │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  SkillEngine v1.0.0                                      │   │
│  │  3 Workflow Prompts + 3 Knowledge Resources              │   │
│  │  Available via SSE native + REST JSON-RPC                │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  Definition Normalizer                                    │   │
│  │  Auto-fixes: run_after→runAfter, default_value→defaultValue│  │
│  │  Injects missing $schema + contentVersion                 │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  Flow Adopter Pipeline (NEW in v3.4.0)                    │   │
│  │  Registers flows in Dataverse workflows entity            │   │
│  │  Enables AddSolutionComponent for non-solution-aware flows│   │
│  │  Equivalent to portal "Add existing → Outside solutions"  │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                 │
│  Transports: SSE (/sse) + REST JSON-RPC (/mcp)                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## Tools (27)

### Environments
| Tool | Description |
|------|-------------|
| `pa-list-environments` | List Power Platform environments (may return empty — see Known Limitations) |

### Flows
| Tool | Description |
|------|-------------|
| `pa-list-flows` | List flows in an environment |
| `pa-get-flow-details` | Get full flow details including definition |
| `pa-create-flow` | Create a cloud flow (auto-normalizes definition). Optional `solutionUniqueName` param for auto-adoption into a solution. |
| `pa-update-flow` | Update flow definition, name, or state (auto-normalizes) |
| `pa-enable-flow` | Enable (start) a flow |
| `pa-disable-flow` | Disable (stop) a flow |
| `pa-delete-flow` | Permanently delete a flow |
| `pa-trigger-flow` | Manually trigger a flow |
| `pa-adopt-flow` | **NEW** — Adopt a non-solution-aware flow into a Dataverse solution. Creates the required Dataverse `workflows` record then calls `AddSolutionComponent`. |

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

## Flow Adopter Pipeline (v3.4.0)

Flows created via the Power Automate portal UI or the Flow Management API are **non-solution-aware** — they exist only in the Flow service, not in Dataverse. The `AddSolutionComponent` action requires a Dataverse `workflows` entity record to exist.

The Flow Adopter solves this by creating the missing record programmatically.

### How It Works

```
pa-adopt-flow (flowId, solutionUniqueName)
  │
  ├─ Step 1: GET flow details from Flow API
  ├─ Step 2: Check Dataverse workflows entity for existing record
  │          └─ If missing → POST /workflows (category=5, type=1)
  └─ Step 3: POST /AddSolutionComponent (componentType=29)
```

### Two Ways to Use It

**Option A — Standalone tool** (for existing flows):
```json
{
  "tool": "pa-adopt-flow",
  "params": {
    "flowId": "b65abe28-4e71-4ff2-8c71-c448cec077a6",
    "solutionUniqueName": "BeakSolution"
  }
}
```

**Option B — Auto-adoption during creation** (new in `pa-create-flow`):
```json
{
  "tool": "pa-create-flow",
  "params": {
    "displayName": "My New Flow",
    "definition": { ... },
    "solutionUniqueName": "BeakSolution"
  }
}
```
When `solutionUniqueName` is provided, `pa-create-flow` automatically:
1. Creates the flow via Flow API
2. Registers it in Dataverse `workflows` entity
3. Calls `AddSolutionComponent` to link it to the solution
4. Returns `_solutionAdoption` status in the response

If auto-adoption fails, the flow creation still succeeds — retry with `pa-adopt-flow`.

### What It Fixes

| Before v3.4.0 | After v3.4.0 |
|--------------|-------------|
| `pa-add-solution-component` → 404 "does not exist" | `pa-adopt-flow` → creates Dataverse record → `AddSolutionComponent` succeeds |
| Manual portal step: Solutions → Add existing → Outside solutions | Fully automated via API |
| Flows created via `pa-create-flow` not solution-aware | Pass `solutionUniqueName` for automatic adoption |

---

## Skills — SkillEngine v1.0.0

Skills are **embedded operational knowledge** that agents read BEFORE calling tools. Available via both SSE and REST transports.

### Workflow Prompts (3)

| Prompt | Purpose | Prevents |
|--------|---------|----------|
| `workflow-auth` | Device Code Flow sequence + environment discovery note | Auth errors, empty environment confusion |
| `workflow-create-flow` | Mandatory 5-step procedure: Auth → Create → Wait → Verify → Add to Solution | Empty definitions, `run_after` errors, missing `$schema`, 404 solution errors |
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
| `default_value` → `defaultValue` | `$connections` parameter block | `pa-create-flow`, `pa-update-flow` |
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
| `pa-list-flows` indexing delay | New flows may take 15–30 min to appear in list results | Use the flow ID from `pa-create-flow` directly — don't search |
| Non-solution-aware flows 404 on AddSolutionComponent | Flows created outside solutions lack Dataverse records | Use `pa-adopt-flow` or pass `solutionUniqueName` to `pa-create-flow` |

---

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | ✅ | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | ✅ | App registration client ID |
| `AZURE_CLIENT_SECRET` | ✅ | App registration client secret |
| `POWER_PLATFORM_ENVIRONMENT_ID` | ✅ | Target Power Platform environment ID |
| `DATAVERSE_URL` | ✅ | Dataverse instance URL (e.g., `bolthousefreshprod.crm.dynamics.com`) |
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
├── index.ts                        # Main entry: Express + MCP + 27 tools + skills
├── auth/
│   ├── azure-token-manager.ts      # Service principal token lifecycle
│   └── user-auth-manager.ts        # Per-user Device Code Flow
├── api/
│   └── power-platform-client.ts    # HTTP client for Flow + PowerApps APIs
├── clients/
│   └── solution-client.ts          # Dataverse Web API client
├── config/
│   └── environment-resolver.ts     # Environment ID resolution
├── tools/
│   └── tool-descriptions.ts        # Centralized tool descriptions
├── utils/
│   ├── param-resolver.ts           # snake_case → camelCase param unwrapper
│   ├── normalize-definition.ts     # Flow definition auto-fixer
│   └── flow-adopter.ts             # Dataverse workflow record creator (v3.4.0)
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
| v3.4.0 | 2026-03-25 | Flow Adopter pipeline (`pa-adopt-flow`), `pa-create-flow` auto-adoption via `solutionUniqueName`, 27 tools, `workflow-create-flow` updated to 5-step procedure |
| v3.3.0 | 2026-03-25 | SkillEngine v1.0.0, REST skill endpoints, definition normalizer, production env switch, param-resolver, pa-create-solution, pa-delete-solution |
| v3.2.0 | 2026-03-22 | 5 Dataverse solution tools |
| v3.1.0 | 2026-03-21 | pa-get-connection, pa-delete-connection |
| v3.0.0 | 2026-03-20 | Dual-token architecture, Device Code Flow, unified SSE + REST |
| v2.4.0 | 2026-03-17 | Service principal auth, SSE transport |

---

## License

Internal — Bolthouse Fresh Foods
