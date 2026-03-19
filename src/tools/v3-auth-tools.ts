/**
 * v3 Auth Tools — ToolDefinition-compatible auth tools
 * 
 * These match the EXACT pattern used in src/index.ts toolDefs array:
 *   { name, description, inputSchema, handler: async (p) => ok({...}) | fail(msg) }
 * 
 * Usage in index.ts:
 *   import { createV3AuthTools } from './tools/v3-auth-tools.js';
 *   import { UserAuthManager } from './auth/user-auth-manager.js';
 *   const userAuthManager = new UserAuthManager();
 *   const toolDefs: ToolDefinition[] = [ ...existingTools, ...createV3AuthTools(userAuthManager) ];
 * 
 * @version 3.0.0
 */

import { ToolResult, ToolDefinition } from '../types.js';
import { UserAuthManager } from '../auth/user-auth-manager.js';

// Match the ok/fail helpers from index.ts
function ok(data: any): ToolResult {
  return { content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }] };
}
function fail(msg: string): ToolResult {
  return { content: [{ type: 'text' as const, text: JSON.stringify({ error: msg }) }], isError: true };
}

/**
 * Creates the 3 auth tool definitions for per-user delegated auth.
 * Returns ToolDefinition[] that can be spread into the toolDefs array in index.ts.
 */
export function createV3AuthTools(userAuthManager: UserAuthManager): ToolDefinition[] {
  return [
    // ---- Auth Tool 1: pa-auth-start ----
    {
      name: 'pa-auth-start',
      description:
        'Start user authentication for Power Automate write operations. ' +
        'Initiates a Device Code Flow — returns a user code and URL. ' +
        'The user must visit the URL and enter the code to complete login. ' +
        'Required before create/update/delete flow operations.',
      inputSchema: {
        type: 'object',
        properties: {
          user_id: {
            type: 'string',
            description:
              'Email or identifier of the user authenticating ' +
              '(e.g., "marie@bolthousefresh.com"). This determines who owns created flows.',
          },
        },
        required: ['user_id'],
      },
      handler: async (p: any) => {
        try {
          if (!p.user_id) return fail('user_id is required');
          const result = await userAuthManager.startAuth(p.user_id);
          return ok({
            status: 'device_code_issued',
            user_id: p.user_id,
            user_code: result.userCode,
            verification_uri: result.verificationUri,
            expires_in_seconds: result.expiresIn,
            instructions: result.message,
            next_step:
              'After the user completes login at the URL above, call pa-auth-poll to finish authentication.',
          });
        } catch (e: any) {
          return fail(e.message);
        }
      },
    },

    // ---- Auth Tool 2: pa-auth-poll ----
    {
      name: 'pa-auth-poll',
      description:
        'Poll for completion of a pending device code authentication. ' +
        'Call this after pa-auth-start once the user has entered their code. ' +
        'Returns "authenticated" on success, "pending" if still waiting, or "expired" if timed out.',
      inputSchema: {
        type: 'object',
        properties: {
          user_id: {
            type: 'string',
            description:
              'Email of the user being authenticated. ' +
              'Optional if only one auth is pending.',
          },
        },
      },
      handler: async (p: any) => {
        try {
          const result = await userAuthManager.pollAuth(p.user_id);
          const response: Record<string, any> = {
            status: result.status,
            message: result.message,
          };
          if (result.userId) response.user_id = result.userId;

          switch (result.status) {
            case 'authenticated':
              response.next_step =
                'Authentication complete. Write operations (create/update/delete flows) are now available.';
              break;
            case 'pending':
              response.next_step =
                'User has not completed login yet. Wait a few seconds and call pa-auth-poll again.';
              break;
            case 'expired':
              response.next_step =
                'Device code expired. Call pa-auth-start to begin a new authentication.';
              break;
            case 'error':
              response.next_step =
                'An error occurred. Review the message and try pa-auth-start again if needed.';
              break;
          }
          return ok(response);
        } catch (e: any) {
          return fail(e.message);
        }
      },
    },

    // ---- Auth Tool 3: pa-auth-status ----
    {
      name: 'pa-auth-status',
      description:
        'Check the current authentication status for Power Automate user operations. ' +
        'Shows whether a user has a valid delegated token, when it expires, ' +
        'and lists all authenticated users if no user_id is specified.',
      inputSchema: {
        type: 'object',
        properties: {
          user_id: {
            type: 'string',
            description:
              'Email of the user to check status for. ' +
              'If omitted, returns status for the most recent or all authenticated users.',
          },
        },
      },
      handler: async (p: any) => {
        try {
          const result = userAuthManager.getStatus(p.user_id);
          const response: Record<string, any> = {
            authenticated: result.authenticated,
            message: result.message,
            delegated_auth_configured: userAuthManager.isConfigured(),
          };
          if (result.userId) response.user_id = result.userId;
          if (result.expiresIn !== undefined) {
            response.expires_in_seconds = result.expiresIn;
            response.expires_in_minutes = Math.floor(result.expiresIn / 60);
          }
          const users = userAuthManager.listAuthenticatedUsers();
          if (users.length > 0) response.authenticated_users = users;
          return ok(response);
        } catch (e: any) {
          return fail(e.message);
        }
      },
    },
  ];
}
