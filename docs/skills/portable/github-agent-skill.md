# GitHub Agent Skill

> **Version:** 1.0.0 | **Format:** Portable Agent Skill
> **Usage:** Add to any AI agent's system prompt or knowledge base.
> **Purpose:** Prevents common GitHub MCP tool failures — especially the SHA trap.

---

## CRITICAL: Tool Selection

You have access to GitHub MCP tools. Follow these rules exactly.

### Writing Files (Create or Update)

**ALWAYS** use `push_files`. **NEVER** use `create_or_update_file`.

```
Tool: push_files
Params:
  owner: <repo owner>
  repo: <repo name>
  branch: <target branch>
  files: [{"path": "path/to/file.ts", "content": "<full file content>"}]
  message: "<conventional commit message>"
```

Why: `create_or_update_file` requires the SHA hash of the existing file.
If you don't have it, the call fails (422). If you fetch it first, it can
go stale before you write (409 conflict). `push_files` needs NO SHA.
It handles create AND update in one atomic call.

### Reading Files

Use `get_file_contents` to read file content or check if a file exists.
Do NOT use `create_or_update_file` for reads.

### Multiple Files = One Call

When changing 2+ files, push them ALL in one `push_files` call:

```
push_files(
  owner: "org",
  repo: "my-repo",
  branch: "main",
  files: [
    {"path": "src/index.ts", "content": "..."},
    {"path": "src/auth.ts", "content": "..."},
    {"path": "package.json", "content": "..."}
  ],
  message: "feat: add auth module"
)
```

This creates ONE commit. Never push related files separately —
it creates broken intermediate states if the repo has CI/CD.

---

## Pre-Push Checklist

Before every `push_files` call, verify:

1. **Interface alignment** — If file A calls methods from file B,
   confirm the method names and signatures match exactly.
   Wrong: `manager.startDeviceCodeFlow()`
   Right: `manager.startAuth()`

2. **Import paths** — Verify paths are correct and include file
   extensions if the project uses ESM (`.js` for TypeScript ESM).

3. **Version sync** — If multiple files declare a version string,
   update ALL of them in the same commit.

4. **Cross-file references** — If file A references fields that
   file B produces, confirm file B actually returns those fields.

---

## Commit Messages

Use conventional commit format:

```
<type>: <concise summary>

Types:
  feat:     New feature
  fix:      Bug fix
  docs:     Documentation
  chore:    Maintenance
  refactor: Restructure (no behavior change)

Examples:
  feat: add per-user OAuth Device Code Flow
  fix: correct method name alignment in auth handler
  docs: update README with deployment instructions
```

---

## Error Recovery

If a push fails:

| Error | Cause | Fix |
|-------|-------|-----|
| 422 "SHA required" | Used `create_or_update_file` | Switch to `push_files` |
| 409 "Conflict" | Stale SHA or concurrent push | Use `push_files` (no SHA needed) |
| 404 "Not found" | Wrong owner, repo, or branch | Verify repo exists and branch name |
| 422 "Invalid path" | Path contains invalid characters | Check for spaces, special chars |

If `push_files` itself fails, check:
- Is the `branch` correct? (case-sensitive)
- Is the `content` a string? (not an object)
- Are all `path` values relative? (no leading `/`)

---

## Branch Strategy

- **Hotfix / small change (1-3 files):** Push directly to `main`
- **New feature (4+ files):** Create feature branch, push there,
  then create a Pull Request for review
- **Always confirm** the target branch with the user if unsure

---

## Quick Rules

```
1. WRITE  → push_files (always)
2. READ   → get_file_contents
3. SHA    → never needed (push_files bypasses it)
4. Multi  → one push_files call (atomic)
5. Commit → conventional format (feat|fix|docs: summary)
6. Verify → check response has ref + sha (confirms success)
```
