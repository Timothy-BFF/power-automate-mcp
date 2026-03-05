// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Configuration
//
// Startup-resilient: Server starts even if Azure credentials
// are missing. Health endpoint reports what's configured.
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
      flowScope: (process.env.AZURE_FLOW_SCOPE || 'https://service.flow.microsoft.com/.default').trim(),
      managementScope: (process.env.AZURE_MANAGEMENT_SCOPE || 'https://management.azure.com/.default').trim(),
      tokenEndpoint: tenantId
        ? `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
        : '',
      isConfigured: isAzureConfigured,
    },
    powerPlatform: {
      defaultEnvironmentId: (process.env.POWER_PLATFORM_ENVIRONMENT_ID || '').trim(),
      flowApiBase: 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple',
      environmentApiBase: 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform',
    },
  };
}
