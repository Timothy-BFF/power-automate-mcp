// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Configuration
//
// ALL values configurable via environment variables.
// No hardcoded URLs, secrets, or infrastructure details.
// Startup-resilient: Server starts even if Azure credentials
// are missing. Health endpoint reports what's configured.
//
// Author: GROW by Bolthouse Fresh (Architected by MCA)
// ═══════════════════════════════════════════════════════════════

import dotenv from 'dotenv';
dotenv.config();

export interface AppConfig {
  port: number;
  logLevel: string;
  simtheoryAuthToken: string;
  azure: {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    flowScope: string;
    managementScope: string;
    tokenEndpoint: string;
    isConfigured: boolean;
  };
  powerPlatform: {
    defaultEnvironmentId: string;
    flowApiBase: string;
    environmentApiBase: string;
  };
}

export function loadConfig(): AppConfig {
  const tenantId = (process.env.AZURE_TENANT_ID || '').trim();
  const clientId = (process.env.AZURE_CLIENT_ID || '').trim();
  const clientSecret = (process.env.AZURE_CLIENT_SECRET || '').trim();

  // Determine if Azure AD is fully configured
  const isAzureConfigured = !!(tenantId && clientId && clientSecret);

  return {
    port: parseInt(process.env.PORT || '8080', 10),
    logLevel: (process.env.LOG_LEVEL || 'info').trim(),
    simtheoryAuthToken: (process.env.SIMTHEORY_AUTH_TOKEN || '').trim(),
    azure: {
      tenantId,
      clientId,
      clientSecret,

      // Flow API scope — for flow CRUD, run history, triggers
      flowScope: (process.env.AZURE_FLOW_SCOPE || 'https://service.flow.microsoft.com/.default').trim(),

      // PowerApps/BAP API scope — for environment listing
      // CRITICAL: The BAP API (api.bap.microsoft.com) requires tokens
      // scoped to service.powerapps.com, NOT management.azure.com.
      // Using management.azure.com scope returns 403 Forbidden.
      managementScope: (process.env.AZURE_MANAGEMENT_SCOPE || 'https://service.powerapps.com/.default').trim(),

      tokenEndpoint: tenantId
        ? `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
        : '',
      isConfigured: isAzureConfigured,
    },
    powerPlatform: {
      defaultEnvironmentId: (process.env.POWER_PLATFORM_ENVIRONMENT_ID || '').trim(),

      // All API base URLs configurable via env vars — no hardcoded infrastructure
      flowApiBase: (process.env.FLOW_API_BASE || 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple').trim(),
      environmentApiBase: (process.env.ENVIRONMENT_API_BASE || 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform').trim(),
    },
  };
}
