/**
 * Resolves Power Platform environment ID with fallback chain:
 * 1. Explicit environmentId parameter (from tool call)
 * 2. POWER_PLATFORM_ENVIRONMENT_ID env var (Railway config)
 * 3. DEFAULT_ENVIRONMENT_ID env var (legacy fallback)
 * 4. Constructed Default-{tenantId} (last resort)
 */
export function resolveEnvironmentId(environmentId?: string): string {
  if (environmentId) {
    console.log(`[EnvResolver] Using explicit parameter: ${environmentId}`);
    return environmentId.trim();
  }

  // Check Railway variable name first
  const ppEnv = (process.env.POWER_PLATFORM_ENVIRONMENT_ID || '').trim();
  if (ppEnv) {
    console.log(`[EnvResolver] Using POWER_PLATFORM_ENVIRONMENT_ID: ${ppEnv}`);
    return ppEnv;
  }

  // Legacy fallback
  const defaultEnv = (process.env.DEFAULT_ENVIRONMENT_ID || '').trim();
  if (defaultEnv) {
    console.log(`[EnvResolver] Using DEFAULT_ENVIRONMENT_ID: ${defaultEnv}`);
    return defaultEnv;
  }

  // Last resort: construct from tenant ID
  const tenantId = (process.env.AZURE_TENANT_ID || '').trim();
  if (tenantId) {
    const constructed = `Default-${tenantId}`;
    console.log(`[EnvResolver] Falling back to constructed: ${constructed}`);
    return constructed;
  }

  throw new Error('No environment ID provided and no default configured. Set POWER_PLATFORM_ENVIRONMENT_ID or AZURE_TENANT_ID.');
}
