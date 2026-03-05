import 'dotenv/config';

export const config = {
  tenantId: (process.env.AZURE_TENANT_ID || '').trim(),
  clientId: (process.env.AZURE_CLIENT_ID || '').trim(),
  clientSecret: (process.env.AZURE_CLIENT_SECRET || '').trim(),
  simtheoryToken: (process.env.SIMTHEORY_AUTH_TOKEN || '').trim(),
  defaultEnvironmentId: (process.env.DEFAULT_ENVIRONMENT_ID || '').trim(),
  port: parseInt(process.env.PORT || '8080', 10),
  logLevel: (process.env.LOG_LEVEL || 'info').trim(),
};

if (!config.defaultEnvironmentId && config.tenantId) {
  config.defaultEnvironmentId = `Default-${config.tenantId}`;
}
