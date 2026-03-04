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
  const tenantId = process.env.AZURE_TENANT_ID || '';
  const clientId = process.env.AZURE_CLIENT_ID || '';
  const clientSecret = process.env.AZURE_CLIENT_SECRET || '';

  // Determine if Azure AD is fully configured
  const isAzureConfigured = !!(tenantId && clientId && clientSecret);

  return {
    port: parseInt(process.env.PORT || '8080', 10),
    logLevel: process.env.LOG_LEVEL || 'info',
    simtheoryAuthToken: process.env.SIMTHEORY_AUTH_TOKEN || '',
    azure: {
      tenantId,
      clientId,
      clientSecret,
      flowScope: process.env.AZURE_FLOW_SCOPE || 'https://service.flow.microsoft.com/.default',
      managementScope: process.env.AZURE_MANAGEMENT_SCOPE || 'https://management.azure.com/.default',
      tokenEndpoint: tenantId
        ? `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
        : '',
      isConfigured: isAzureConfigured,
    },
    powerPlatform: {
      defaultEnvironmentId: process.env.POWER_PLATFORM_ENVIRONMENT_ID || '',
      flowApiBase: 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple',
      environmentApiBase: 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform',
    },
  };
}
