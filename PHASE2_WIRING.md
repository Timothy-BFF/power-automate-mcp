# Phase 2 Wiring Guide — v3.0.0 Integration

This guide documents the exact changes needed in existing files to complete
the v3.0.0 per-user delegated auth integration.

## Overview

Phase 1 created:
- `src/auth/user-auth-manager.ts` — Device Code Flow
- `src/auth/token-router.ts` — Dual-token dispatcher
- `src/auth/index.ts` — Barrel exports
- `src/config/delegated-auth-settings.ts` — New env vars
- `src/config/auth-tool-definitions.ts` — Auth tool schemas

Phase 2 created:
- `src/tools/auth-tool-handlers.ts` — Handler implementations
- `src/v3-bootstrap.ts` — Singleton initialization module
- This wiring guide

## Changes Required in Existing Files

### 1. `src/index.ts` — Main Server Entry Point

**Add these imports** (near the top, with other imports):

```typescript
import { initV3Auth, v3 } from './v3-bootstrap';
```

**Add initialization** (after AzureTokenManager is created, before server starts):

```typescript
// Initialize v3 dual-token auth system
// Pass the existing token manager's getToken method
initV3Auth(() => azureTokenManager.getToken());
```

**Add auth tools to tools/list** (in the ListToolsRequestSchema handler):

```typescript
// In the tools/list handler, spread the auth tools into the existing array:
tools: [
  ...existingToolDefinitions,
  ...v3.authToolDefinitions,  // ← ADD THIS LINE
]
```

**Add auth tool dispatch to tools/call** (at the TOP of the CallToolRequestSchema handler):

```typescript
// At the start of the tools/call handler, before the existing switch/if:
const authResult = await v3.dispatchAuthTool(toolName, args);
if (authResult) return authResult;

// ... existing tool handling continues below ...
```

### 2. `src/api/power-platform-client.ts` — API Client

The existing client likely calls `this.tokenManager.getToken()` to get
the service principal token for all requests. For v3, write operations
need to use the user's delegated token instead.

**Option A: Minimal Change (Recommended)**

Add an optional `tokenOverride` parameter to methods that perform writes:

```typescript
// In each write method (createFlow, updateFlow, deleteFlow, etc.):
async createFlow(
  environmentId: string,
  definition: any,
  tokenOverride?: string  // ← ADD THIS
): Promise<any> {
  const token = tokenOverride || await this.tokenManager.getToken();
  // ... rest of method uses `token` for the Authorization header
}
```

Then in `index.ts`, before calling write methods:

```typescript
// For write operations, get the user token via TokenRouter:
const token = await v3.getToken('pa-create-flow', userId);
const result = await powerPlatformClient.createFlow(envId, definition, token);
```

**Option B: Full Integration**

Replace the token source entirely:

```typescript
// Replace:
const token = await this.tokenManager.getToken();

// With:
const token = await v3.getToken(toolName, userId);
```

This requires passing `toolName` and `userId` through the call chain.

### 3. `src/config/settings.ts` — Environment Variables

**Add the new env vars** to whatever config object/function loads them:

```typescript
// Add to the settings/config:
PA_USER_CLIENT_ID: process.env.PA_USER_CLIENT_ID || '',
PA_USER_TENANT_ID: process.env.PA_USER_TENANT_ID || process.env.AZURE_TENANT_ID || '',
```

Note: The `UserAuthManager` reads these directly from `process.env`,
so this step is optional but recommended for documentation/validation.

## Railway Environment Variables

Add these to the `intelligent-youthfulness` Railway project:

```
PA_USER_CLIENT_ID=<your-device-code-app-registration-client-id>
PA_USER_TENANT_ID=<your-tenant-id>  (or omit if same as AZURE_TENANT_ID)
```

## Azure AD App Registration Requirements

The `PA_USER_CLIENT_ID` app registration must have:

1. **Authentication** → Allow public client flows → **Yes** (required for Device Code Flow)
2. **API Permissions** → `https://service.flow.microsoft.com/.default` (delegated)
3. **Supported account types** → Accounts in this organizational directory only

## Testing

### Smoke Test: Auth Tools

1. Call `pa-auth-status` → should return `{ authenticated: false, delegated_auth_configured: true }`
2. Call `pa-auth-start` with `{ user_id: "timothy@bolthousefresh.com" }` → should return device code + URL
3. Visit the URL, enter the code, sign in
4. Call `pa-auth-poll` → should return `{ status: "authenticated" }`
5. Call `pa-auth-status` → should show authenticated with token expiry

### Smoke Test: Write Operations

1. Authenticate via steps above
2. Call `pa-create-flow` → should use the user's delegated token
3. Verify the flow appears in Power Automate portal under the user's account

### Regression Test: Read Operations

1. Call `pa-list-flows` → should still work via service principal (no user auth needed)
2. Call `pa-get-flow-details` → same
3. Call `pa-get-run-history` → same
