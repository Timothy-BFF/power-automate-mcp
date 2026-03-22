# GitHub Operations Procedure for AI Agents

> **Version:** 1.0.0 | **Last Updated:** 2026-03-22
>
> **Purpose:** Mandatory operational rules for AI agents performing GitHub
> operations via MCP tools. These rules prevent the recurring failures
> observed across multiple MCA development sessions.

---

## RULE 1: ALWAYS Use `push_files` — NEVER `create_or_update_file`

This is the single most important rule.

### The SHA Trap

The `create_or_update_file` tool requires the **SHA hash** of the existing
file before it can update it. This creates a fragile two-step process:

```
# BAD: Two-call round trip, SHA can go stale
Step 1: GET file → extract SHA
Step 2: PUT file with SHA → hope nobody pushed in between
```

If the SHA is stale (another commit happened between GET and PUT), the
update **fails silently or with a 409 conflict**.

### The Correct Pattern

```
# GOOD: Single atomic operation, no SHA needed
push_files(owner, repo, branch, files=[
  { path: "src/index.ts", content: "..." },
  { path: "src/auth/user-auth-manager.ts", content: "..." }
], message: "feat: add user auth manager")
```

### Why `push_files` Is Superior

| Feature | `create_or_update_file` | `push_files` |
|---------|------------------------|-------------|
| SHA required | YES (must GET first) | NO |
| Multi-file atomic | NO (one file per call) | YES (all files in one commit) |
| Race condition risk | HIGH (stale SHA) | NONE |
| API calls per push | 2N (GET + PUT per file) | 1 (single call) |
| Commit history | N commits for N files | 1 clean commit |

### When To Use Each

- **`push_files`** — ALWAYS. For creating, updating, or replacing files.
- **`create_or_update_file`** — NEVER for updates. Only acceptable for
  reading file content (when `get_file_contents` is unavailable).
- **`get_file_contents`** — Use for READING files or checking if they exist.

---

## RULE 2: Pre-Push Validation Checklist

Before EVERY `push_files` call, mentally verify:

### Interface Alignment Check

```
[ ] Method names match the actual class/interface
    BAD:  userAuthManager.startDeviceCodeFlow()
    GOOD: userAuthManager.startAuth()

[ ] Return field names match the actual interface
    BAD:  result.user_code, result.verification_uri
    GOOD: result.userCode, result.verificationUri

[ ] Import paths are correct
    BAD:  import { UserAuthManager } from './user-auth-manager'
    GOOD: import { UserAuthManager } from './auth/user-auth-manager.js'
```

### Version Sync Check

```
[ ] Banner/constant version matches across files
    - src/index.ts: const VERSION = '3.0.3'
    - package.json: "version": "3.0.3"  ← MUST MATCH
    - README.md version references ← SHOULD MATCH

[ ] If bumping version, bump ALL locations in the same commit
```

### Cross-File Reference Check

```
[ ] If file A references fields from file B, verify B actually produces them
    Example: tool-descriptions.ts references _fetchedVia
             → power-platform-client.ts MUST produce _fetchedVia

[ ] If skill docs reference response fields, verify the handler returns them
    Example: flow-creation-procedure.md says check _definitionStatus
             → getFlowDetails() MUST include _definitionStatus in response
```

---

## RULE 3: Atomic Multi-File Commits

When a change spans multiple files, push them ALL in one commit.

### BAD: Sequential Single-File Pushes

```
# This creates a broken intermediate state on Railway
push_files(files=[{path: 'src/index.ts', content: '...'}], message: 'update index')
# Railway auto-deploys HERE → build fails because auth-tool-handlers.ts
# still has old method names
push_files(files=[{path: 'src/tools/auth-tool-handlers.ts', content: '...'}], message: 'update handlers')
```

### GOOD: Single Atomic Commit

```
push_files(files=[
  { path: 'src/index.ts', content: '...' },
  { path: 'src/tools/auth-tool-handlers.ts', content: '...' },
  { path: 'src/api/power-platform-client.ts', content: '...' }
], message: 'fix: align method names across all files')
# Railway auto-deploys ONCE → all files consistent → build succeeds
```

---

## RULE 4: Commit Message Format

Use conventional commits. This is not optional.

```
FORMAT: <type>: <summary> (v<version>)

TYPES:
  feat:  New feature or capability
  fix:   Bug fix or correction
  docs:  Documentation only
  chore: Maintenance, deps, config
  refactor: Code restructure, no behavior change

EXAMPLES:
  feat: add UserAuthManager with Device Code Flow (v3.0.3)
  fix: align method names with UserAuthManager interface (v3.0.3)
  docs: add github-operations-procedure.md skill (v1.0.0)
  chore: update package.json version to 3.0.3
```

### Commit Body (for complex changes)

When pushing 3+ files or making architectural changes, include a body:

```
fix: add _fetchedVia metadata to getFlowDetails (v3.0.3)

Implements dual-path GET in PowerPlatformClient:
  - Delegated user endpoint (full definition)
  - Admin endpoint fallback (shell only)

New response fields:
  - _fetchedVia: 'delegated (user@...)' or 'admin (service-principal)'
  - _definitionStatus: 'POPULATED' or 'EMPTY_OR_NOT_RETURNED'
  - _definitionNote: explanation for agent decision-making

Aligns with: tool-descriptions.ts, flow-creation-procedure.md
```

---

## RULE 5: Repository Structure Cache

Do NOT re-discover the repo structure every session. Use this cache.

### power-automate-mcp (Timothy-BFF/power-automate-mcp)

```
Language: TypeScript | Deploy: Railway | Branch: main

src/
├── index.ts                          ← Entry point, tool registration, Express app
├── types.ts                          ← Shared type definitions
├── auth/
│   ├── azure-token-manager.ts        ← Service principal token (client credentials)
│   └── user-auth-manager.ts          ← Per-user delegated token (Device Code Flow)
├── api/
│   └── power-platform-client.ts      ← HTTP client (dual-token, all API calls)
├── config/
│   └── environment-resolver.ts       ← Resolves Power Platform environment ID
├── tools/
│   ├── tool-descriptions.ts          ← All tool descriptions (single source of truth)
│   └── auth-tool-handlers.ts         ← pa-auth-start/poll/status (SSE registration)
docs/
└── skills/
    ├── flow-creation-procedure.md    ← Flow creation guidance for agents
    └── github-operations-procedure.md ← THIS FILE
```

### power-interpreter (BolthouseFreshFoods/power-interpreter)

```
Language: Python | Deploy: Railway | Branch: main

app/
├── main.py                           ← FastAPI entry, MCP handler
├── auth.py                           ← API key auth
├── config.py                         ← Env vars, blocked builtins
├── mcp_server.py                     ← MCP tool registration
├── data_manager.py                   ← Session/file management
├── database.py                       ← PostgreSQL connection
├── fetch_from_url.py                 ← URL fetching tool
├── models.py                         ← SQLAlchemy models
├── response_guard.py                 ← Smart response truncation
├── engine/
│   ├── executor.py                   ← Code execution sandbox
│   └── skill_engine.py               ← Skill registration & dispatch
├── microsoft/
│   ├── auth_manager.py               ← OAuth Device Code Flow (per-user)
│   ├── graph_client.py               ← Microsoft Graph HTTP client
│   └── tools.py                      ← ms_auth, onedrive, sharepoint tools
├── routes/
│   └── ...                           ← API routes
docs/
└── ...
patches/
└── ...
```

---

## RULE 6: Post-Push Verification Protocol

After every push, verify deployment health:

### Step 1: Confirm Push Success

The `push_files` response includes a `ref` and `url`. Verify:
- `ref` matches the target branch (e.g., `refs/heads/main`)
- No error in the response

### Step 2: Wait for Railway Deploy

Railway auto-deploys on push to `main`. Typical timeline:
- Build start: ~10-30s after push
- Build complete: ~60-90s
- Container start: ~5-10s after build

### Step 3: Parse Deploy Logs (Structured)

When user provides Railway logs, extract:

| Check | What to Look For | Healthy Value |
|-------|------------------|---------------|
| Version | `Power Automate MCP v{X.Y.Z}` | Matches pushed version |
| Tools | `MCP tools registered: N` | Expected count (13 MCP + 3 auth = 16) |
| Auth | `Dual-token mode` | Present if UserAuth configured |
| Tokens | `Token acquired for scope:` | BAP + Flow + PowerApps (3 lines) |
| Errors | Any `severity: error` from app service | ZERO |
| Port | `running on port 8080` | Must match Railway PORT |

### Step 4: If Errors Found

```
IF TypeScript compilation error (TS2339, TS2345, etc.):
  → Interface mismatch. Check RULE 2 alignment.
  → Fix and push ATOMIC commit with all affected files.

IF Module not found:
  → Check import paths (.js extension required in ESM TypeScript)
  → Verify file exists at the path referenced

IF Token acquisition failed:
  → Check Railway environment variables are set
  → AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
```

---

## RULE 7: Reading Files Before Editing

When you need to modify an existing file:

### Preferred: Full File Rewrite via push_files

```
# If you know the full intended content:
push_files(files=[{ path: 'src/index.ts', content: '<full file content>' }])
```

### When You Need Current Content First

```
# Use get_file_contents to READ, then push_files to WRITE
get_file_contents(owner, repo, path='src/index.ts')
# Parse the returned content
# Modify as needed
# Push the complete modified file via push_files
```

**NEVER** use `create_or_update_file` for the write step.
The GET is only to read current content — the write is ALWAYS `push_files`.

---

## RULE 8: Branch Strategy

### Direct to Main (Default for Hotfixes)

Use when:
- Fixing a build-breaking error (TS compilation failures)
- Single-file doc updates
- Version bumps
- The change is < 3 files and well-understood

```
push_files(branch='main', files=[...], message='fix: ...')
```

### Feature Branch + PR (For New Features)

Use when:
- Adding a new capability (e.g., new MCP tool)
- Architectural changes (e.g., adding UserAuthManager)
- Changes spanning 5+ files
- User explicitly requests a PR for review

```
# Step 1: Create branch
create_branch(branch='feature/user-auth', from_branch='main')

# Step 2: Push to feature branch
push_files(branch='feature/user-auth', files=[...], message='feat: ...')

# Step 3: Create PR
create_pull_request(title='feat: Add UserAuthManager', head='feature/user-auth', base='main')

# Step 4: User reviews → merge
```

---

## Quick Reference Card

```
┌─────────────────────────────────────────────────────────┐
│                  GITHUB OPERATIONS                       │
│                  Quick Reference                         │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  WRITE files  → push_files (ALWAYS)                     │
│  READ files   → get_file_contents                       │
│  SHA needed?  → NEVER (push_files bypasses it)          │
│                                                         │
│  Multi-file?  → ONE push_files call (atomic)            │
│  Hotfix?      → Direct to main                          │
│  New feature? → Feature branch + PR                     │
│                                                         │
│  Pre-push:    → Interface check ✓                       │
│              → Version sync ✓                           │
│              → Cross-file refs ✓                        │
│                                                         │
│  Post-push:   → Confirm ref in response                 │
│              → Wait for Railway deploy                  │
│              → Parse logs (structured table)             │
│                                                         │
│  Commit msg:  → feat|fix|docs|chore: summary (vX.Y.Z)  │
│                                                         │
└─────────────────────────────────────────────────────────┘
```
