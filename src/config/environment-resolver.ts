import { config } from './index.js';

export function resolveEnvironmentId(envId?: string): string {
  return envId || config.defaultEnvironmentId;
}
