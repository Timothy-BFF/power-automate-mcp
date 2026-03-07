# ID Bridge Resolution — Integration Guide (v2.4.0)

> **Author:** MCA (Model Context Architect)  
> **Date:** 2026-03-07  
> **Evidence:** MCA Smoke Test v2.3.3  
> **Branch:** `fix/id-bridge-v2.4.0`  
> **Module:** `src/api/id-bridge.ts`

---

## Problem Statement

When `pa-create-flow` creates a flow via the Dataverse `workflows` entity, the POST returns a `workflowid` (primary key). This ID is a **Dataverse GUID** — the Flow Admin API (`/providers/Microsoft.ProcessSimple/...`) does NOT recognize it.

The Flow Admin API uses a *different* GUID stored in the Dataverse `workflowidunique` field, which is populated **asynchronously** after creation (5–30s propagation delay, confirmed by IT).

### Evidence from Smoke Test (2026-03-07)

| Event | ID Returned | Works with Flow Admin? |
|---|---|---|
| `create-flow` (1st call) | `e45ae471-658e-4278-ac8b-5d8db6cc138b` | ❌ 404 (×6 attempts) |
| `list-flows` (same flow) | `e45ae471-658e-4278-ac86-5b7a3ba03519` | ✅ |
| Shared GUID prefix | `e45ae471-658e-4278-ac8` (23 chars) | — |
| `_idMapping.dataverseWorkflowId` | `434bf8b5-b919-f111-8342-000d3a150318` | — |

The `get-flow-details` call was attempted **6 consecutive times** with the Dataverse GUID. All returned `FlowNotFound` 404.

---

## Architecture

```
┌──────────────┐     POST /workflows     ┌──────────────────┐
│  MCP Client  │ ──────────────────────→ │    Dataverse     │
│ (SimTheory)  │ ← workflowid (PK)      │   workflows()    │
└──────┬───────┘                         └───────┬──────────┘
       │                                         │
       │  NEW: resolveAfterCreate()              │ async propagation
       │  polls for workflowidunique             │ (5–30 seconds)
       │                                         ▼
       │                                ┌──────────────────────┐
       │  GET /flows/{resolvedId}       │  workflowidunique     │
       │ ─────────────────────────────→ │  field populated      │
       │                                └──────────────────────┘
       │                                         │
       │                                         ▼
       │                                ┌──────────────────────┐
       │ ← flow details + _idMapping   │  Flow Admin API      │
       │                                │  /flows/{id}         │
       └────────────────────────────────└──────────────────────┘
```

### Bug Coverage

| Bug | Symptom | Fix Function |
|---|---|---|
| #1 | `create-flow` returns Dataverse GUID → 404 on follow-up calls | `resolveAfterCreate()` |
| #2 | No wait for `workflowidunique` propagation | `resolveAfterCreate()` (polling loop) |
| #3 | `get-flow-details` returns 404 with no fallback | `getFlowWithFallback()` |

---

## Integration Steps

### Step 1: Import the module

In `src/api/power-platform-client.ts`:

```typescript
import {
  resolveAfterCreate,
  getFlowWithFallback,
  type IdBridgeDeps,
  type CreateFlowResolution,
} from './id-bridge.js';
```

### Step 2: Create the dependency adapter

Add a private method (or getter) to your client class that builds the `IdBridgeDeps` object. This adapter bridges the existing API client methods to the interface expected by id-bridge.

```typescript
private get idBridgeDeps(): IdBridgeDeps {
  return {
    getWorkflowRecord: async (workflowId, select) => {
      // Use your existing Dataverse GET method.
      // The URL should be:
      //   GET {dataverseUrl}/api/data/v9.2/workflows({workflowId})?$select={select}
      try {
        return await this.dataverseGet(`workflows(${workflowId})`, { $select: select });
      } catch (err: any) {
        if (err?.status === 404 || err?.message?.includes('404')) return null;
        throw err;
      }
    },

    getFlow: async (flowId, envId) => {
      // Use your existing Flow Admin GET method.
      // The URL should be:
      //   GET .../environments/{envId}/flows/{flowId}?api-version=2016-11-01
      try {
        return await this.getFlowRaw(flowId, envId);
      } catch (err: any) {
        if (err?.status === 404 || err?.message?.includes('404')) return null;
        throw err;
      }
    },

    listFlows: async (envId, top = 50) => {
      // Use your existing list-flows method.
      const result = await this.listFlowsRaw(envId, 'all', top);
      return result?.value ?? result ?? [];
    },
  };
}
```

> **Note:** The method names (`dataverseGet`, `getFlowRaw`, `listFlowsRaw`) are
> placeholders — adapt to your actual method names. The key contract:
> - `getWorkflowRecord` returns `null` on 404, throws on other errors
> - `getFlow` returns `null` on 404, throws on other errors
> - `listFlows` returns an array of flow summary objects

### Step 3: Wire into createFlow (Bug #1 + #2)

Find your `createFlow` method. After the Dataverse POST succeeds and you have the Dataverse primary key:

**BEFORE (current behavior — returns wrong ID):**
```typescript
// After Dataverse POST:
const dataverseId = response.workflowid; // or from OData-EntityId header
return {
  status: 'created',
  flowId: dataverseId,  // ← BUG: This is the Dataverse GUID, not the Flow API ID
  displayName,
  state: state ?? 'Stopped',
};
```

**AFTER (correct behavior — resolved ID with metadata):**
```typescript
// After Dataverse POST:
const dataverseId = response.workflowid; // or from OData-EntityId header

// NEW — Wait for propagation and resolve the Flow API ID
const resolution = await resolveAfterCreate(
  dataverseId,
  envId,
  this.idBridgeDeps,
  displayName,  // enables list-flows fallback
);

return {
  status: 'created',
  flowId: resolution.flowId,             // ← FIXED: Flow API ID
  displayName,
  state: state ?? 'Stopped',
  _idMapping: resolution._idMapping,     // ← NEW: Resolution metadata
  _propagation: resolution._propagation, // ← NEW: Propagation timing
};
```

### Step 4: Wire into getFlowDetails (Bug #3)

Find your `getFlowDetails` method.

**BEFORE (current — no fallback on 404):**
```typescript
async getFlowDetails(flowId: string, envId: string) {
  const url = `.../${envId}/flows/${flowId}?api-version=2016-11-01`;
  const resp = await this.authenticatedFetch(url);
  if (resp.status === 404) {
    throw new Error(`API Error 404: Could not find flow '${flowId}'.`);
  }
  return resp.json();
}
```

**AFTER (with fallback cascade):**
```typescript
async getFlowDetails(flowId: string, envId: string) {
  const result = await getFlowWithFallback(flowId, envId, this.idBridgeDeps);
  return {
    ...result.data,
    _idMapping: result._idMapping,  // include resolution metadata
  };
}
```

### Step 5: Bump version

In `package.json`:
```json
"version": "2.4.0"
```

---

## Configuration

The default propagation config is:

```typescript
{
  maxAttempts: 6,      // 6 poll iterations
  baseDelayMs: 3_000,  // 3s initial delay
  backoff: 1.5,        // 1.5× multiplier
  maxDelayMs: 15_000,  // 15s max per poll
}
```

**Resulting poll schedule:**
| Attempt | Delay | Cumulative |
|---|---|---|
| 1 | 3.0s | 3.0s |
| 2 | 4.5s | 7.5s |
| 3 | 6.75s | 14.25s |
| 4 | 10.1s | 24.35s |
| 5 | 15.0s (capped) | 39.35s |
| 6 | 15.0s (capped) | 54.35s |

**Worst case:** ~54 seconds if Dataverse never resolves, then falls back to list-flows.
This covers the observed 5–30 second propagation window with margin.

To override, pass a partial config:
```typescript
const resolution = await resolveAfterCreate(
  dataverseId, envId, this.idBridgeDeps, displayName,
  { maxAttempts: 8, baseDelayMs: 2000 },  // custom config
);
```

---

## Testing

### Smoke Test Sequence

1. **Create a test flow:**
   ```
   pa-create-flow "ID Bridge Test v2.4.0" { definition... } Stopped
   ```
   **Expected response fields:**
   - `flowId`: A Flow API ID (NOT the Dataverse GUID)
   - `_idMapping.resolvedVia`: `"workflowidunique"`
   - `_propagation.verified`: `true`
   - `_propagation.delayMs`: Typically 3000–15000
   - `_propagation.attempts`: Typically 1–3

2. **Get flow details with the returned ID:**
   ```
   pa-get-flow-details {flowId from step 1}
   ```
   **Expected:** Returns full flow details (NOT a 404)

3. **Get flow details with a Dataverse GUID (regression test):**
   ```
   pa-get-flow-details {_idMapping.dataverseWorkflowId from step 1}
   ```
   **Expected:** Resolves via fallback, returns flow details with
   `_idMapping.resolvedVia: "workflowidunique"`

4. **Clean up:**
   ```
   pa-delete-flow {flowId from step 1}
   ```

### Expected Logs — Success Path (Bug #1 + #2)

```
[IdBridge] Resolving Flow API ID for Dataverse ID: {dvId}
[IdBridge] Propagation poll 1/6 — waiting 3000ms
[IdBridge] Candidate: {flowApiId} (via workflowidunique). Verifying on Flow Admin API...
[IdBridge] ✓ Propagation verified in 3200ms (1 attempts, workflowidunique)
```

### Expected Logs — Fallback Path (Bug #3)

```
[IdBridge] Direct lookup: {dvGuid}
[IdBridge] 404 on direct lookup for {dvGuid}. Starting fallback resolution...
[IdBridge] Attempting Dataverse resolution for {dvGuid}...
[IdBridge] Dataverse resolved: {dvGuid} → {flowApiId} (via workflowidunique)
```

### Expected Logs — Worst Case (Degraded)

```
[IdBridge] Resolving Flow API ID for Dataverse ID: {dvId}
[IdBridge] Propagation poll 1/6 — waiting 3000ms
[IdBridge] Poll 1: workflowidunique not yet populated
[IdBridge] Propagation poll 2/6 — waiting 4500ms
...
[IdBridge] Propagation poll 6/6 — waiting 15000ms
[IdBridge] Dataverse poll exhausted after 6 attempts. Searching list-flows for "..."
[IdBridge] ✓ Resolved via list-flows search: {flowApiId} (42000ms)
```

---

## Rollback

If issues occur after deployment:

1. **Revert the import** and `resolveAfterCreate` / `getFlowWithFallback` calls in `power-platform-client.ts`
2. **`id-bridge.ts` can remain** in the codebase — it has zero side effects and no global state
3. Redeploy

The module is pure functions with dependency injection. It touches nothing unless explicitly called.

---

## Future Improvements (v2.5.0+)

| # | Improvement | Description |
|---|---|---|
| 4 | Cache EnvResolver | `[EnvResolver]` re-runs on every request — should cache at startup |
| 5 | Structured error logs | Error bodies logged line-by-line; should be single structured entry |
| 6 | Suppress npm startup noise | Empty log lines from `npm run start` |
| 7 | Request correlation IDs | No way to trace concurrent requests through logs |
| 8 | Idempotency detection logging | No log distinguishing "created new" vs "returning existing" |
