/**
 * Auth module barrel export
 * 
 * v2.4.0: AzureTokenManager, SimtheoryAuth, TokenRefreshMiddleware
 * v3.0.0: + UserAuthManager, TokenRouter, AuthRequiredError
 */

export { UserAuthManager, AuthRequiredError } from './user-auth-manager';
export type { AuthStartResult, AuthPollResult, AuthStatusResult } from './user-auth-manager';
export { TokenRouter } from './token-router';
export type { OperationType } from './token-router';
