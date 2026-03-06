/**
 * Resolves Power Platform environment ID with fallback chain:
 * 1. Explicit environmentId parameter (from tool call)
 * 2. POWER_PLATFORM_ENVIRONMENT_ID env var (Railway config)
 * 3. DEFAULT_ENVIRONMENT_ID env var (legacy fallback)
 * 4. Constructed Default-{tenantId} (last resort)
 */
export function resolveEnvironmentId(environmentId?: string): string {
  if (environmentId) return environmentId.trim();

  // Check Railway variable name first
  const ppEnv = (process.env.POWER_PLATFORM_ENVIRONMENT_ID || '').trim();
  if (ppEnv) return ppEnv;

  // Legacy fallback
  const defaultEnv = (process.env.DEFAULT_ENVIRONMENT_ID || '').trim();
  if (defaultEnv) return defaultEnv;

  // Last resort: construct from tenant ID
  const tenantId = (process.env.AZURE_TENANT_ID || '').trim();
  if (tenantId) return `Default-${tenantId}`;

  throw new Error('No environment ID provided and no default configured. Set POWER_PLATFORM_ENVIRONMENT_ID or AZURE_TENANT_ID.');
}
