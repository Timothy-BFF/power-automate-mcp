// ═══════════════════════════════════════════════════════════════
// Power Automate MCP Server — Configuration
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
  };
  powerPlatform: {
    defaultEnvironmentId: string;
    flowApiBase: string;
    environmentApiBase: string;
  };
}

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export function loadConfig(): AppConfig {
  const tenantId = requireEnv('AZURE_TENANT_ID');

  return {
    port: parseInt(process.env.PORT || '8080', 10),
    logLevel: process.env.LOG_LEVEL || 'info',
    simtheoryAuthToken: requireEnv('SIMTHEORY_AUTH_TOKEN'),
    azure: {
      tenantId,
      clientId: requireEnv('AZURE_CLIENT_ID'),
      clientSecret: requireEnv('AZURE_CLIENT_SECRET'),
      flowScope: process.env.AZURE_FLOW_SCOPE || 'https://service.flow.microsoft.com/.default',
      managementScope: process.env.AZURE_MANAGEMENT_SCOPE || 'https://management.azure.com/.default',
      tokenEndpoint: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    },
    powerPlatform: {
      defaultEnvironmentId: process.env.POWER_PLATFORM_ENVIRONMENT_ID || '',
      flowApiBase: 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple',
      environmentApiBase: 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform',
    },
  };
}
