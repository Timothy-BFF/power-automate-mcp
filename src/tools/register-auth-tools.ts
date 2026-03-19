/**
 * Register auth tools with McpServer for SSE transport.
 *
 * Root cause: The v3 auth tools were defined as ToolDefinition objects
 * for the REST JSON-RPC toolDefs array, but never registered with the
 * McpServer SDK. Simtheory connects via SSE, which uses McpServer —
 * so it only saw 12 tools instead of 15.
 *
 * This file registers the same 3 auth tools using the server.tool()
 * pattern (identical to pa-list-connections), making them visible
 * to SSE clients.
 *
 * Usage in index.ts:
 *   import { registerAuthTools } from './tools/register-auth-tools.js';
 *   registerAuthTools(mcpServer, userAuthManager);
 *
 * @version 3.0.1
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { UserAuthManager } from '../auth/user-auth-manager.js';

export function registerAuthTools(server: McpServer, userAuthManager: UserAuthManager): void {

  // ---- pa-auth-start ----
  (server as any).tool(
    'pa-auth-start',
    'Start user authentication for Power Automate write operations. ' +
    'Initiates a Device Code Flow — returns a user code and URL. ' +
    'The user must visit the URL and enter the code to complete login. ' +
    'Required before create/update/delete flow operations.',
    {
      user_id: z.string().describe(
        'Email or identifier of the user authenticating ' +
        '(e.g., "marie@bolthousefresh.com"). This determines who owns created flows.'
      ),
    },
    async (args: any) => {
      try {
        if (!args.user_id) {
          return {
            content: [{ type: 'text' as const, text: JSON.stringify({ error: 'user_id is required' }) }],
            isError: true,
          };
        }
        const result = await userAuthManager.startAuth(args.user_id);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              status: 'device_code_issued',
              user_id: args.user_id,
              user_code: result.userCode,
              verification_uri: result.verificationUri,
              expires_in_seconds: result.expiresIn,
              instructions: result.message,
              next_step: 'After the user completes login at the URL above, call pa-auth-poll to finish authentication.',
            }, null, 2),
          }],
        };
      } catch (e: any) {
        return {
          content: [{ type: 'text' as const, text: JSON.stringify({ error: e.message }) }],
          isError: true,
        };
      }
    }
  );

  // ---- pa-auth-poll ----
  (server as any).tool(
    'pa-auth-poll',
    'Poll for completion of a pending device code authentication. ' +
    'Call this after pa-auth-start once the user has entered their code. ' +
    'Returns "authenticated" on success, "pending" if still waiting, or "expired" if timed out.',
    {
      user_id: z.string().optional().describe(
        'Email of the user being authenticated. Optional if only one auth is pending.'
      ),
    },
    async (args: any) => {
      try {
        const result = await userAuthManager.pollAuth(args.user_id);
        const response: Record<string, any> = {
          status: result.status,
          message: result.message,
        };
        if (result.userId) response.user_id = result.userId;

        switch (result.status) {
          case 'authenticated':
            response.next_step = 'Authentication complete. Write operations (create/update/delete flows) are now available.';
            break;
          case 'pending':
            response.next_step = 'User has not completed login yet. Wait a few seconds and call pa-auth-poll again.';
            break;
          case 'expired':
            response.next_step = 'Device code expired. Call pa-auth-start to begin a new authentication.';
            break;
          case 'error':
            response.next_step = 'An error occurred. Review the message and try pa-auth-start again if needed.';
            break;
        }

        return {
          content: [{ type: 'text' as const, text: JSON.stringify(response, null, 2) }],
        };
      } catch (e: any) {
        return {
          content: [{ type: 'text' as const, text: JSON.stringify({ error: e.message }) }],
          isError: true,
        };
      }
    }
  );

  // ---- pa-auth-status ----
  (server as any).tool(
    'pa-auth-status',
    'Check the current authentication status for Power Automate user operations. ' +
    'Shows whether a user has a valid delegated token, when it expires, ' +
    'and lists all authenticated users if no user_id is specified.',
    {
      user_id: z.string().optional().describe(
        'Email of the user to check status for. Optional — shows all authenticated users if omitted.'
      ),
    },
    async (args: any) => {
      try {
        const result = userAuthManager.getStatus(args.user_id);
        const response: Record<string, any> = {
          authenticated: result.authenticated,
          message: result.message,
        };
        if (result.userId) response.user_id = result.userId;
        if (result.expiresIn !== undefined) response.token_expires_in_seconds = result.expiresIn;
        if (!result.authenticated) {
          response.authenticated_users = userAuthManager.listAuthenticatedUsers();
        }
        return {
          content: [{ type: 'text' as const, text: JSON.stringify(response, null, 2) }],
        };
      } catch (e: any) {
        return {
          content: [{ type: 'text' as const, text: JSON.stringify({ error: e.message }) }],
          isError: true,
        };
      }
    }
  );

  console.log('[AuthTools] 3 auth tools registered with McpServer (SSE transport)');
}
