# Phase 2 Wiring Guide — v3.0.0 Integration

Now that I've seen the actual source code, here are the EXACT changes needed.

## Architecture Insight

- `src/config/` contains only `environment-resolver.ts` and `index.ts` (no separate settings or tool-definitions files)
- Tool definitions are inline in `src/index.ts` via `const toolDefs: ToolDefinition[]`
- Each tool uses `{ name, description, inputSchema, handler }` pattern
- Handlers use `ok(data)` and `fail(msg)` helpers
- `PowerPlatformClient` uses `flowAdminRequest()` for all Flow API calls

---

## Change 1: `src/index.ts` — Register Auth Tools (4 lines)

### Step 1a: Add imports (after existing imports, around line 9)

```typescript
// === v3.0.0: Per-User Delegated Auth ===
import { UserAuthManager } from './auth/user-auth-manager.js';
import { createV3AuthTools } from './tools/v3-auth-tools.js';
```

### Step 1b: Initialize UserAuthManager (after `const client = new PowerPlatformClient(tokenManager);`)

```typescript
// === v3.0.0: Initialize per-user auth ===
const userAuthManager = new UserAuthManager();
```

### Step 1c: Add auth tools to toolDefs array

Find the `const toolDefs: ToolDefinition[] = [` line and change the array to spread in auth tools:

```typescript
const toolDefs: ToolDefinition[] = [
  // ... all existing tool definitions stay exactly as they are ...

  // === v3.0.0: Per-User Auth Tools ===
  ...createV3AuthTools(userAuthManager),
];
```

Alternatively, add after the closing of the last existing tool entry:

```typescript
  // (end of last existing tool)
  },
  // === v3.0.0: Per-User Auth Tools (pa-auth-start, pa-auth-poll, pa-auth-status) ===
  ...createV3AuthTools(userAuthManager),
];
```

That's it. **4 lines total** to add auth tools.

---

## Change 2: `src/api/power-platform-client.ts` — Token Override for Writes

This enables write operations to use the user's delegated token instead of
the service principal token.

### Step 2a: Add optional `tokenOverride` to `flowAdminRequest`

Find the `flowAdminRequest` method (private method that all Flow API calls go through).
It likely looks like:

```typescript
private async flowAdminRequest(path: string, method: string = 'GET', body?: any): Promise<any> {
  const token = await this.tokenManager.getToken('https://service.flow.microsoft.com/.default');
  // ... axios/fetch with Authorization: Bearer ${token}
}
```

Add an optional 4th parameter:

```typescript
private async flowAdminRequest(
  path: string,
  method: string = 'GET',
  body?: any,
  tokenOverride?: string  // ← ADD THIS
): Promise<any> {
  const token = tokenOverride || await this.tokenManager.getToken('https://service.flow.microsoft.com/.default');
  // ... rest stays the same
}
```

### Step 2b: Thread `tokenOverride` through write methods

For each write method, add `tokenOverride` parameter and pass it through:

```typescript
async createFlow(envId: string, ..., tokenOverride?: string): Promise<any> {
  // ... existing body construction ...
  const result = await this.flowAdminRequest(
    `/providers/Microsoft.ProcessSimple/scopes/admin/environments/${envId}/flows?api-version=${FLOW_API_VER}`,
    'POST',
    body,
    tokenOverride  // ← ADD THIS
  );
  // ... rest stays the same
}

async updateFlow(envId: string, flowId: string, updates: {...}, tokenOverride?: string): Promise<any> {
  // ... pass tokenOverride to flowAdminRequest
}

async deleteFlow(envId: string, flowId: string, tokenOverride?: string): Promise<any> {
  return this.flowAdminRequest(
    `/providers/.../${flowId}?api-version=${FLOW_API_VER}`,
    'DELETE',
    undefined,
    tokenOverride  // ← ADD THIS
  );
}

async enableDisableFlow(envId: string, flowId: string, action: 'start' | 'stop', tokenOverride?: string): Promise<any> {
  return this.flowAdminRequest(
    `/providers/.../${flowId}/${action}?api-version=${FLOW_API_VER}`,
    'POST',
    undefined,
    tokenOverride  // ← ADD THIS
  );
}

async triggerFlow(envId: string, flowId: string, body?: any, tokenOverride?: string): Promise<any> {
  return this.flowAdminRequest(
    `/providers/.../${flowId}/triggers/manual/run?api-version=${FLOW_API_VER}`,
    'POST',
    body || {},
    tokenOverride  // ← ADD THIS
  );
}

async cancelRun(envId: string, flowId: string, runId: string, tokenOverride?: string): Promise<any> {
  return this.flowAdminRequest(
    `/providers/.../${runId}/cancel?api-version=${FLOW_API_VER}`,
    'POST',
    undefined,
    tokenOverride  // ← ADD THIS
  );
}
```

### Step 2c: Wire token override in index.ts write tool handlers

In each WRITE tool handler in `index.ts`, get the user token before calling the client:

```typescript
// Example for pa-create-flow handler:
handler: async (p: any) => {
  try {
    const envId = resolveEnvironmentId(p.environmentId);
    // === v3.0.0: Get user token for write operations ===
    let userToken: string | undefined;
    try {
      userToken = await userAuthManager.getAccessToken() || undefined;
    } catch { /* falls back to service principal */ }
    const r = await client.createFlow(envId, p.displayName, p.definition, ..., userToken);
    return ok(r);
  } catch (e: any) { return fail(e.message); }
},
```

---

## Change 3: Version Bump

In `src/index.ts`, update:
```typescript
const VERSION = '3.0.0';
```

---

## Railway Environment Variables

Add to `intelligent-youthfulness` project on Railway:

```
PA_USER_CLIENT_ID=<your-device-code-app-registration-client-id>
PA_USER_TENANT_ID=<your-tenant-id>
```

---

## Azure AD App Registration Checklist

The `PA_USER_CLIENT_ID` app must have:

1. **Authentication → Allow public client flows → Yes** (required for Device Code Flow)
2. **API Permissions → Flow Service: `https://service.flow.microsoft.com/.default`** (delegated)
3. **Supported account types → Single tenant** (this org directory only)

---

## Phased Rollout

### Phase 2a (Auth Tools Only — no existing code changes)
Just add the 4 lines in Change 1 to `index.ts`. This gives you:
- `pa-auth-start` / `pa-auth-poll` / `pa-auth-status` tools
- Users can authenticate
- Write operations still use service principal (existing behavior)

### Phase 2b (Token Override — minimal existing code changes)
Apply Change 2 to `power-platform-client.ts` and the write tool handlers.
- Write operations now use user's delegated token when available
- Falls back to service principal if user isn't authenticated

Recommendation: Ship Phase 2a first, test auth flow, then Phase 2b.

---

## Smoke Test

1. Deploy to Railway
2. Call `pa-auth-status` → `{ authenticated: false, delegated_auth_configured: true }`
3. Call `pa-auth-start { user_id: "timothy@bolthousefresh.com" }` → device code
4. Visit URL, enter code, sign in
5. Call `pa-auth-poll` → `{ status: "authenticated" }`
6. Call `pa-auth-status` → shows token expiry
7. Call `pa-list-flows` → still works (service principal, no change)
8. (Phase 2b) Call `pa-create-flow` → uses Timothy's delegated token
