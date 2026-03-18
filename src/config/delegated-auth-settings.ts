/**
 * Delegated Auth Settings — v3.0.0 Addition
 * 
 * Configuration for per-user Device Code Flow authentication.
 * Uses a SEPARATE Azure AD App Registration from the Service Principal.
 * 
 * Environment Variables:
 *   PA_USER_CLIENT_ID   — App Registration for delegated (device code) auth
 *   PA_USER_TENANT_ID   — Tenant ID (falls back to AZURE_TENANT_ID)
 */

export interface DelegatedAuthConfig {
  clientId: string;
  tenantId: string;
  isConfigured: boolean;
}

export function getDelegatedAuthConfig(): DelegatedAuthConfig {
  const clientId = process.env.PA_USER_CLIENT_ID || '';
  const tenantId = process.env.PA_USER_TENANT_ID || process.env.AZURE_TENANT_ID || '';

  return {
    clientId,
    tenantId,
    isConfigured: !!(clientId && tenantId),
  };
}

/**
 * Validate delegated auth configuration at startup.
 * Logs warnings if not configured (non-fatal — service principal still works for reads).
 */
export function validateDelegatedAuthConfig(): void {
  const config = getDelegatedAuthConfig();

  if (!config.isConfigured) {
    console.warn('╔══════════════════════════════════════════════════════════════╗');
    console.warn('║  [v3.0.0] Delegated Auth NOT configured                     ║');
    console.warn('║  Write operations will fail until PA_USER_CLIENT_ID is set   ║');
    console.warn('║  Read operations will continue via Service Principal         ║');
    console.warn('╚══════════════════════════════════════════════════════════════╝');
  } else {
    console.log('[DelegatedAuthConfig] ✓ Configured');
    console.log(`  Tenant: ${config.tenantId}`);
    console.log(`  Client: ${config.clientId.substring(0, 8)}...`);
  }
}
