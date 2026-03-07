// ════════════════════════════════════════════════════════════════
// ID Bridge Resolution — Power Automate MCP v2.4.0
// ════════════════════════════════════════════════════════════════
//
// Resolves Dataverse workflow GUIDs → Flow Admin API IDs.
//
// Fixes:
//   Bug #1: create-flow returns Dataverse PK instead of Flow API ID
//   Bug #2: No post-creation propagation wait
//   Bug #3: get-flow-details has no 404 → resolution fallback
//
// Usage: Import and wire into power-platform-client.ts
//   See docs/ID-BRIDGE-v2.4.0.md for integration steps.
//
// Evidence: MCA Smoke Test v2.3.3 (2026-03-07)
//   create-flow returned: e45ae471-658e-4278-ac8b-5d8db6cc138b  (Dataverse GUID → 404)
//   list-flows showed:    e45ae471-658e-4278-ac86-5b7a3ba03519  (Flow API ID → works)
//   6 consecutive 404s on get-flow-details with the Dataverse GUID.
// ════════════════════════════════════════════════════════════════

// ─── Types ──────────────────────────────────────────────────────

export interface IdBridgeConfig {
  /** Max propagation poll iterations (default: 6) */
  maxAttempts: number;
  /** Initial delay between polls in ms (default: 3000) */
  baseDelayMs: number;
  /** Delay multiplier per attempt (default: 1.5) */
  backoff: number;
  /** Maximum delay cap per attempt in ms (default: 15000) */
  maxDelayMs: number;
}

export interface IdMapping {
  /** The resolved Flow Admin API ID (usable with /flows/{id}) */
  flowApiId: string;
  /** The Dataverse workflow primary key */
  dataverseWorkflowId: string;
  /** How the ID was resolved */
  resolvedVia: 'workflowidunique' | 'name' | 'list-flows' | 'direct';
}

export interface PropagationMeta {
  /** Whether the resolved ID was verified on the Flow Admin API */
  verified: boolean;
  /** Total elapsed time for resolution in ms */
  delayMs: number;
  /** Number of poll/resolution attempts made */
  attempts: number;
}

export interface CreateFlowResolution {
  /** The correct Flow API ID to return to callers */
  flowId: string;
  /** Mapping metadata showing Dataverse → Flow API translation */
  _idMapping: IdMapping;
  /** Propagation wait metadata */
  _propagation: PropagationMeta;
}

/**
 * Dependency contract — injected from the existing API client.
 * This keeps id-bridge.ts decoupled from auth, HTTP clients, and URL construction.
 */
export interface IdBridgeDeps {
  /**
   * Query a single Dataverse workflow record by its primary key.
   * Should call: GET {dataverseUrl}/api/data/v9.2/workflows({workflowId})?$select={select}
   * @returns The record object, or null if not found.
   */
  getWorkflowRecord(
    workflowId: string,
    select: string,
  ): Promise<Record<string, any> | null>;

  /**
   * GET a single flow from the Flow Admin API.
   * Should call: GET .../environments/{envId}/flows/{flowId}?api-version=2016-11-01
   * @returns The flow object, or null on 404. Should throw on non-404 errors.
   */
  getFlow(
    flowId: string,
    envId: string,
  ): Promise<Record<string, any> | null>;

  /**
   * List flows from the Flow Admin API.
   * Should call: GET .../environments/{envId}/v2/flows?$filter=all&$top={top}
   * @returns Array of flow summary objects.
   */
  listFlows(envId: string, top?: number): Promise<Array<Record<string, any>>>;
}

// ─── Defaults ───────────────────────────────────────────────────

const DEFAULTS: IdBridgeConfig = {
  maxAttempts: 6,
  baseDelayMs: 3_000,
  backoff: 1.5,
  maxDelayMs: 15_000,
};

// ─── Internal helpers ───────────────────────────────────────────

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

function log(level: 'info' | 'warn' | 'error', msg: string): void {
  const fn =
    level === 'error'
      ? console.error
      : level === 'warn'
        ? console.warn
        : console.log;
  fn(`[IdBridge] ${msg}`);
}

function is404(err: unknown): boolean {
  if (err == null || typeof err !== 'object') return false;
  const e = err as Record<string, any>;
  return (
    e.status === 404 ||
    e.statusCode === 404 ||
    e.response?.status === 404 ||
    String(e.message ?? '').includes('404') ||
    String(e.message ?? '').includes('FlowNotFound')
  );
}

function buildResult(
  flowApiId: string,
  dvId: string,
  via: IdMapping['resolvedVia'],
  delayMs: number,
  attempts: number,
  verified: boolean,
): CreateFlowResolution {
  return {
    flowId: flowApiId,
    _idMapping: { flowApiId, dataverseWorkflowId: dvId, resolvedVia: via },
    _propagation: { verified, delayMs, attempts },
  };
}

// ════════════════════════════════════════════════════════════════
// Bug #1 + #2: Post-Creation ID Resolution with Propagation Wait
// ════════════════════════════════════════════════════════════════
//
// After a Dataverse POST creates a workflow, the Flow Admin API
// doesn't immediately recognize the new flow. The correct Flow
// API ID lives in the `workflowidunique` field, which is populated
// asynchronously (5-30s observed, confirmed by IT).
//
// This function polls Dataverse until the field is populated,
// then verifies the ID works on the Flow Admin API.
// ════════════════════════════════════════════════════════════════

/**
 * Resolve the correct Flow API ID after a Dataverse workflow creation.
 *
 * Strategy chain:
 *   1. Poll Dataverse for `workflowidunique` (primary strategy)
 *   2. Search `list-flows` by display name (fallback)
 *   3. Return unverified Dataverse ID with warning (degraded)
 *
 * @param dataverseId  - The ID returned from the Dataverse POST (workflowid PK or similar)
 * @param envId        - Power Platform environment ID
 * @param deps         - Injected API dependencies
 * @param displayName  - Flow display name (enables list-flows fallback)
 * @param config       - Optional config overrides
 */
export async function resolveAfterCreate(
  dataverseId: string,
  envId: string,
  deps: IdBridgeDeps,
  displayName?: string,
  config: Partial<IdBridgeConfig> = {},
): Promise<CreateFlowResolution> {
  const cfg = { ...DEFAULTS, ...config };
  const t0 = Date.now();
  let attempts = 0;

  log('info', `Resolving Flow API ID for Dataverse ID: ${dataverseId}`);

  // ── Strategy 1: Poll Dataverse for workflowidunique ────────────
  for (let i = 1; i <= cfg.maxAttempts; i++) {
    attempts = i;
    const delay = Math.min(
      Math.round(cfg.baseDelayMs * Math.pow(cfg.backoff, i - 1)),
      cfg.maxDelayMs,
    );

    log('info', `Propagation poll ${i}/${cfg.maxAttempts} — waiting ${delay}ms`);
    await sleep(delay);

    try {
      const rec = await deps.getWorkflowRecord(
        dataverseId,
        'workflowidunique,name,statecode',
      );

      if (!rec) {
        log('warn', `Poll ${i}: Dataverse record not found yet`);
        continue;
      }

      // Extract the candidate Flow API ID
      const candidate = rec.workflowidunique ?? rec.name;

      if (!candidate) {
        log('warn', `Poll ${i}: workflowidunique not yet populated`);
        continue;
      }

      // Don't accept the candidate if it's the same as the input
      // (means the field hasn't resolved to the Flow API ID yet)
      if (candidate === dataverseId) {
        log('warn', `Poll ${i}: workflowidunique matches input — not yet resolved`);
        continue;
      }

      const via: IdMapping['resolvedVia'] = rec.workflowidunique
        ? 'workflowidunique'
        : 'name';

      log('info', `Candidate: ${candidate} (via ${via}). Verifying on Flow Admin API...`);

      // Verify the candidate is visible on the Flow Admin API
      try {
        const flow = await deps.getFlow(candidate, envId);
        if (flow) {
          const elapsed = Date.now() - t0;
          log(
            'info',
            `✓ Propagation verified in ${elapsed}ms (${attempts} attempts, ${via})`,
          );
          return buildResult(candidate, dataverseId, via, elapsed, attempts, true);
        }
        log('warn', `Candidate ${candidate} not yet visible on Flow Admin API`);
      } catch (flowErr) {
        if (!is404(flowErr)) throw flowErr; // non-404 = real error, don't swallow
        log('warn', `Candidate ${candidate}: Flow Admin API returned 404 — not propagated yet`);
      }
    } catch (dvErr) {
      log('warn', `Poll ${i} Dataverse query failed: ${(dvErr as Error).message}`);
    }
  }

  // ── Strategy 2: Search list-flows by display name ──────────────
  if (displayName) {
    log(
      'warn',
      `Dataverse poll exhausted after ${attempts} attempts. ` +
        `Searching list-flows for "${displayName}"...`,
    );

    try {
      const flows = await deps.listFlows(envId, 100);
      const match = flows.find((f) => {
        const name =
          f.properties?.displayName ?? f.displayName ?? f.properties?.definition?.metadata?.name;
        return name === displayName;
      });

      if (match) {
        const fid: string = match.name ?? match.id;
        const elapsed = Date.now() - t0;
        log('info', `✓ Resolved via list-flows search: ${fid} (${elapsed}ms)`);
        return buildResult(fid, dataverseId, 'list-flows', elapsed, attempts + 1, true);
      }

      log('warn', `Flow "${displayName}" not found in list-flows (${flows.length} flows scanned)`);
    } catch (err) {
      log('error', `List-flows fallback failed: ${(err as Error).message}`);
    }
  }

  // ── Strategy 3: Degraded — return Dataverse ID with warning ────
  const elapsed = Date.now() - t0;
  log(
    'error',
    `✗ Resolution FAILED after ${elapsed}ms (${attempts} attempts). ` +
      `Returning unverified Dataverse ID. Caller may get 404s.`,
  );
  return buildResult(dataverseId, dataverseId, 'direct', elapsed, attempts, false);
}

// ════════════════════════════════════════════════════════════════
// Bug #3: get-flow-details with 404 → ID Resolution Fallback
// ════════════════════════════════════════════════════════════════
//
// When get-flow-details receives a 404, the provided ID may be a
// Dataverse GUID. This function cascades through resolution
// strategies to find the correct Flow API ID.
// ════════════════════════════════════════════════════════════════

/**
 * Get flow details with automatic 404 → ID resolution fallback.
 *
 * Strategy chain on 404:
 *   1. Direct lookup (provided ID as-is)
 *   2. Treat ID as Dataverse PK → query for workflowidunique → retry
 *   3. Partial GUID prefix match on list-flows → retry
 *
 * @param flowId - The flow ID (may be a Flow API ID or Dataverse GUID)
 * @param envId  - Power Platform environment ID
 * @param deps   - Injected API dependencies
 */
export async function getFlowWithFallback(
  flowId: string,
  envId: string,
  deps: IdBridgeDeps,
): Promise<{ data: Record<string, any>; _idMapping: IdMapping }> {
  // ── Attempt 1: Direct lookup ──────────────────────────────────
  log('info', `Direct lookup: ${flowId}`);
  try {
    const flow = await deps.getFlow(flowId, envId);
    if (flow) {
      return {
        data: flow,
        _idMapping: {
          flowApiId: flowId,
          dataverseWorkflowId: flowId,
          resolvedVia: 'direct',
        },
      };
    }
  } catch (err: unknown) {
    if (!is404(err)) throw err; // non-404 = real error
    log('warn', `404 on direct lookup for ${flowId}. Starting fallback resolution...`);
  }

  // ── Attempt 2: Dataverse GUID → workflowidunique resolution ───
  try {
    log('info', `Attempting Dataverse resolution for ${flowId}...`);
    const rec = await deps.getWorkflowRecord(flowId, 'workflowidunique,name');

    if (rec) {
      const resolved = rec.workflowidunique ?? rec.name;

      if (resolved && resolved !== flowId) {
        const via: IdMapping['resolvedVia'] = rec.workflowidunique
          ? 'workflowidunique'
          : 'name';
        log('info', `Dataverse resolved: ${flowId} → ${resolved} (via ${via})`);

        try {
          const flow = await deps.getFlow(resolved, envId);
          if (flow) {
            return {
              data: flow,
              _idMapping: {
                flowApiId: resolved,
                dataverseWorkflowId: flowId,
                resolvedVia: via,
              },
            };
          }
        } catch (flowErr) {
          if (!is404(flowErr)) throw flowErr;
          log('warn', `Resolved ID ${resolved} also returned 404`);
        }
      } else {
        log('warn', `Dataverse record found but no distinct workflowidunique`);
      }
    } else {
      log('warn', `No Dataverse record found for ${flowId}`);
    }
  } catch (dvErr) {
    log('warn', `Dataverse resolution failed: ${(dvErr as Error).message}`);
  }

  // ── Attempt 3: Partial GUID prefix match on list-flows ────────
  //
  // Evidence from smoke test: Dataverse GUID and Flow API ID share
  // the first 23 characters, then diverge:
  //   DV:   e45ae471-658e-4278-ac8b-5d8db6cc138b
  //   Flow: e45ae471-658e-4278-ac86-5b7a3ba03519
  //   Match: e45ae471-658e-4278-ac8  (23 chars)
  //
  // We use this as a fuzzy-match heuristic as a last resort.
  log('warn', `Trying partial GUID prefix match on list-flows...`);
  try {
    // Use first 23 chars as the shared prefix (covers groups 1-3 + start of group 4)
    const prefix = flowId.substring(0, 23);
    const flows = await deps.listFlows(envId, 100);

    const match = flows.find((f) => {
      const fid: string = f.name ?? f.id ?? '';
      return fid.startsWith(prefix) && fid !== flowId;
    });

    if (match) {
      const resolved: string = match.name ?? match.id;
      log('info', `Partial GUID match found: ${resolved}`);

      try {
        const flow = await deps.getFlow(resolved, envId);
        if (flow) {
          return {
            data: flow,
            _idMapping: {
              flowApiId: resolved,
              dataverseWorkflowId: flowId,
              resolvedVia: 'list-flows',
            },
          };
        }
      } catch (flowErr) {
        if (!is404(flowErr)) throw flowErr;
        log('error', `Partial GUID match ${resolved} also returned 404`);
      }
    } else {
      log('warn', `No partial GUID match found (scanned ${flows.length} flows, prefix: ${prefix})`);
    }
  } catch (err) {
    log('error', `List-flows search failed: ${(err as Error).message}`);
  }

  // ── All strategies exhausted ──────────────────────────────────
  const errorMsg =
    `FlowNotFound: Could not resolve '${flowId}' via any strategy ` +
    `(direct → 404, Dataverse → no match, list-flows → no match). ` +
    `The ID may be a Dataverse GUID that hasn't propagated, or the flow may not exist.`;
  log('error', errorMsg);
  throw new Error(errorMsg);
}
