# Power Automate Flow Creation Procedure

> **Version:** 1.0.0 | **Last Updated:** 2026-03-22
> **Applies to:** `pa-create-flow`, `pa-update-flow`, `pa-get-flow-details`

## Purpose

This document defines the **mandatory sequence** for creating and verifying
Power Automate flows via the MCP tools. AI agents MUST follow this procedure
to avoid false-negative "empty definition" conclusions and destructive
delete/recreate loops.

## Background

The Power Automate MCP uses a **dual-token architecture**:

| Token Type | Used For | API Path |
|---|---|---|
| Service Principal | Read operations (list, get details) | `/scopes/admin/environments/{env}/...` |
| Delegated User | Write operations (create, update, trigger) | `/environments/{env}/...` |

**Critical behavior:** The admin GET endpoint (`/scopes/admin/`) does NOT return
the `properties.definition` body for flows created via the delegated user endpoint.
It returns only shell metadata (displayName, state, createdTime). This is a
Microsoft API scope limitation, not a bug in the MCP tool.

## Required Sequence

### Step 1: Authenticate (REQUIRED before any write operation)

```
Call: pa-auth-start
  Input: { user_id: "user@company.com" }
  Output: { user_code: "ABCD1234", verification_uri: "https://microsoft.com/devicelogin" }

Instruct user: "Go to microsoft.com/devicelogin and enter code ABCD1234"

Call: pa-auth-poll
  Input: { user_id: "user@company.com" }
  Output: { status: "authenticated" }   <-- WAIT for this before proceeding
```

**DO NOT proceed to Step 2 until `status: "authenticated"` is confirmed.**
Write operations without authentication will return 401 Unauthorized.

### Step 2: Create the Flow (single atomic call)

```
Call: pa-create-flow
  Input: {
    displayName: "My Flow Name",
    definition: {             <-- FULL definition with triggers + actions
      "$schema": "...",
      "triggers": { ... },    <-- At least one trigger required
      "actions": { ... }      <-- At least one action required
    },
    state: "Stopped"
  }
```

The tool sends **shell + definition in ONE API call**. There is no need to
create a shell first and then update it with a definition.

**Success indicators in the response:**
- `status: "created"` -- flow was created
- `_authType: "delegated"` -- used the correct auth path
- `flowId: "xxxxxxxx-xxxx-..."` -- the new flow ID

**If the create returned 200/201, the definition IS saved.** Trust this response.

### Step 3: Wait for Propagation (MANDATORY)

```
Wait at least 10 seconds before calling pa-get-flow-details.
```

Power Automate has a propagation delay between the write endpoint and the
read endpoint. Calling get-flow-details immediately or in parallel with
create-flow will return stale/empty data.

**NEVER call pa-get-flow-details in parallel with pa-create-flow.**

### Step 4: Verify (with same auth session)

```
Call: pa-get-flow-details
  Input: { flowId: "<flowId from Step 2>" }
```

**Check these fields in order:**

1. `_fetchedVia` -- which API path was used?
   - `"delegated (user@...)"` -- full definition returned, results are reliable
   - `"admin (service-principal)"` -- may show empty definition (scope limitation)

2. `_definitionStatus` -- was the definition found?
   - `"POPULATED"` -- triggers and actions are present, flow is confirmed working
   - `"EMPTY_OR_NOT_RETURNED"` -- read `_definitionNote` before taking any action

3. `_definitionNote` -- explains what happened and what to do next

**If `_fetchedVia` shows `admin` and definition appears empty:**
This is the scope mismatch, NOT a missing definition. The definition is saved
but the admin endpoint cannot return it. Re-authenticate and call again, or
ask the user to verify in the Power Automate portal designer.

## Critical Rules

### NEVER DO
- Delete a flow because `pa-get-flow-details` shows empty definition
- Call `pa-get-flow-details` in parallel with `pa-create-flow`
- Call `pa-get-flow-details` within 10 seconds of creating/updating
- Attempt write operations without authenticating first
- Conclude "definition not persisting" from an admin-path GET response
- Create a shell first, then try to add definition in a separate call

### ALWAYS DO
- Authenticate via `pa-auth-start` / `pa-auth-poll` before writes
- Send the full definition (triggers + actions) in the `pa-create-flow` call
- Trust the create/update response -- if it returned 200, the definition IS saved
- Wait 10+ seconds before verifying with `pa-get-flow-details`
- Check `_fetchedVia` before interpreting an empty definition
- If `_fetchedVia` = admin and definition empty, tell user to check portal
- After portal confirmation, proceed to enable the flow

## Common Error Patterns

### Pattern 1: "Empty Definition" False Alarm
```
Agent creates flow -> 200 OK, _authType: delegated
Agent immediately calls get-flow-details -> triggers: [], actions: []
Agent concludes: "definition not saved" -> WRONG
Agent deletes flow and recreates -> destructive loop
```
**Fix:** Wait 10s, check `_fetchedVia`, trust the create response.

### Pattern 2: 401 on Write Operations
```
Agent calls pa-create-flow -> 401 Unauthorized
Agent retries -> 401 again
```
**Fix:** Call `pa-auth-start` / `pa-auth-poll` first. Tokens expire after ~1 hour.

### Pattern 3: Definition Empty After Re-Auth
```
Agent re-authenticates -> pa-auth-poll: authenticated
Agent calls pa-get-flow-details -> _fetchedVia: "delegated (user@...)" but still empty
```
**Fix:** The flow genuinely has no definition. Recreate with full definition.

## Flow Lifecycle Summary

```
[Auth] -> [Create w/ Full Definition] -> [Wait 10s] -> [Verify] -> [Portal Auth] -> [Enable]
```

All steps are sequential. No parallelism. No shortcuts.
