/**
 * Universal Parameter Resolver — Power Automate MCP v3.3.0
 *
 * Problem: Different MCP clients/bridges wrap tool arguments differently:
 *   - Direct: { connectorId: "shared_office365", user_id: "jose@..." }
 *   - Wrapped: { arguments: { connectorId: "..." } }
 *   - Nested: { params: { arguments: { connectorId: "..." } } }
 *   - SDK: { name: "pa-create-connection", arguments: { ... } }
 *
 * This resolver normalizes all patterns into flat key-value pairs.
 * Also handles snake_case/camelCase variations.
 *
 * Usage:
 *   const args = resolveParams(rawParams, ['connectorId', 'user_id', 'environmentId']);
 *   // args.connectorId is guaranteed to be the value or undefined
 */

export function resolveParams(
  raw: any,
  expectedKeys: string[]
): Record<string, any> {
  if (!raw || typeof raw !== 'object') {
    console.warn('[ParamResolver] Received non-object params:', typeof raw, raw);
    return {};
  }

  // Log raw shape for diagnostics
  const topKeys = Object.keys(raw);
  console.log(`[ParamResolver] Raw param keys: [${topKeys.join(', ')}]`);

  // Try unwrapping common nesting patterns
  const candidates: any[] = [
    raw,                          // Direct: { connectorId: "..." }
    raw.arguments,                // MCP standard: { arguments: { connectorId: "..." } }
    raw.params,                   // Alt wrapper: { params: { connectorId: "..." } }
    raw.input,                    // Alt wrapper: { input: { connectorId: "..." } }
    raw.params?.arguments,        // Double-wrapped: { params: { arguments: { ... } } }
    raw.data,                     // Alt wrapper: { data: { connectorId: "..." } }
  ].filter(c => c && typeof c === 'object');

  // Score each candidate by how many expected keys it contains
  let best = raw;
  let bestScore = 0;

  for (const candidate of candidates) {
    let score = 0;
    for (const key of expectedKeys) {
      if (candidate[key] !== undefined) score++;
      // Also check snake_case variant
      const snakeKey = key.replace(/[A-Z]/g, m => '_' + m.toLowerCase());
      if (snakeKey !== key && candidate[snakeKey] !== undefined) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  if (bestScore === 0 && candidates.length > 1) {
    console.warn(`[ParamResolver] No expected keys found in any nesting level. Dumping raw:`);
    console.warn(`[ParamResolver] ${JSON.stringify(raw).substring(0, 500)}`);
  } else if (best !== raw) {
    console.log(`[ParamResolver] Unwrapped params (score: ${bestScore}/${expectedKeys.length})`);
  }

  // Build result with snake_case fallbacks
  const result: Record<string, any> = { ...best };
  for (const key of expectedKeys) {
    if (result[key] === undefined) {
      // Try snake_case variant
      const snakeKey = key.replace(/[A-Z]/g, m => '_' + m.toLowerCase());
      if (snakeKey !== key && result[snakeKey] !== undefined) {
        result[key] = result[snakeKey];
        console.log(`[ParamResolver] Mapped ${snakeKey} -> ${key}`);
      }
    }
  }

  return result;
}
