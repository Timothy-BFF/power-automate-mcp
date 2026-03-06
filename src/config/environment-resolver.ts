export function resolveEnvironmentId(environmentId?: string): string {
  if (environmentId) return environmentId.trim();
  const defaultEnv = (process.env.DEFAULT_ENVIRONMENT_ID || '').trim();
  if (defaultEnv) return defaultEnv;
  const tenantId = (process.env.AZURE_TENANT_ID || '').trim();
  if (tenantId) return `Default-${tenantId}`;
  throw new Error('No environment ID provided and no default configured. Set DEFAULT_ENVIRONMENT_ID or AZURE_TENANT_ID.');
}
