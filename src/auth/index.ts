/**
 * Auth module barrel exports
 *
 * Re-exports all auth managers and types for clean imports.
 * All relative imports use explicit .js extensions per node16 moduleResolution.
 *
 * @version 3.0.1
 */

export { UserAuthManager } from './user-auth-manager.js';
export type { AuthStartResult, AuthPollResult, AuthStatusResult } from './user-auth-manager.js';

export { TokenRouter } from './token-router.js';
export type { TokenRouterConfig } from './token-router.js';
