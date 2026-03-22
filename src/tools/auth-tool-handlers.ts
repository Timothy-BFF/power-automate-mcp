/**
 * Power Automate MCP - Auth Tool Handlers (v3.0.3)
 *
 * Registers the 3 Device Code Flow authentication tools.
 * These MUST be called before any write operation.
 * Descriptions imported from tool-descriptions.ts.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { UserAuthManager } from '../auth/user-auth-manager.js';
import { TOOL_DESCRIPTIONS } from './tool-descriptions.js';

// =========================================================================
// Response Helpers
// =========================================================================

function jsonResponse(data: any): { content: Array<{ type: 'text'; text: string }> } {
  return { content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }] };
}

function errorResponse(error: any): { content: Array<{ type: 'text'; text: string }>; isError: true } {
  const msg = error?.message || String(error);
  console.error(`[AuthTool] Error: ${msg}`);
  return { content: [{ type: 'text' as const, text: `Error: ${msg}` }], isError: true };
}

// =========================================================================
// Auth Tool Registration
// =========================================================================

export function registerAuthTools(
  server: McpServer,
  userAuth: UserAuthManager
): number {
  let count = 0;

  // -----------------------------------------------------------------------
  // pa-auth-start: Initiate Device Code Flow
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-start',
    TOOL_DESCRIPTIONS['pa-auth-start'],
    {
      user_id: z.string().optional().describe('User email (e.g., jose@company.com). If omitted, uses last authenticated user.'),
    },
    async ({ user_id }) => {
      try {
        const result = await userAuth.startDeviceCodeFlow(user_id);
        return jsonResponse({
          status: 'device_code_issued',
          user_code: result.user_code,
          verification_uri: result.verification_uri,
          message: result.message || `To sign in, visit ${result.verification_uri} and enter the code: ${result.user_code}`,
          expires_in: result.expires_in,
          interval: result.interval || 5,
          _note: 'After the user enters the code and signs in, call pa-auth-poll to complete authentication.',
        });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // pa-auth-poll: Poll for authentication completion
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-poll',
    TOOL_DESCRIPTIONS['pa-auth-poll'],
    {
      user_id: z.string().optional().describe('User email to poll for. If omitted, polls for the most recent pending auth.'),
    },
    async ({ user_id }) => {
      try {
        const result = await userAuth.pollForToken(user_id);

        if (result.status === 'authenticated') {
          return jsonResponse({
            status: 'authenticated',
            user_id: result.user_id || user_id || userAuth.getDefaultUserId(),
            message: `Successfully authenticated as ${result.user_id || user_id || userAuth.getDefaultUserId()}. Write operations are now available.`,
            _note: 'You can now use pa-create-flow, pa-update-flow, and pa-trigger-flow.',
          });
        }

        if (result.status === 'pending') {
          return jsonResponse({
            status: 'pending',
            message: 'User has not yet completed login. Call pa-auth-poll again in 5 seconds.',
            _note: 'The user needs to visit microsoft.com/devicelogin and enter the code from pa-auth-start.',
          });
        }

        // Expired or error
        return jsonResponse({
          status: result.status || 'error',
          message: result.message || 'Authentication failed or expired. Call pa-auth-start to try again.',
          error: result.error,
        });
      } catch (error: any) {
        // Handle authorization_pending as a "pending" response, not an error
        if (error.message?.includes('authorization_pending')) {
          return jsonResponse({
            status: 'pending',
            message: 'User has not yet completed login. Call pa-auth-poll again in 5 seconds.',
          });
        }
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // pa-auth-status: Check current authentication state
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-status',
    TOOL_DESCRIPTIONS['pa-auth-status'],
    {
      user_id: z.string().optional().describe('User email to check. If omitted, checks default/last authenticated user.'),
    },
    async ({ user_id }) => {
      try {
        const hasUser = userAuth.hasAuthenticatedUser();
        const defaultUser = userAuth.getDefaultUserId();
        const targetUser = user_id || defaultUser;

        if (!hasUser) {
          return jsonResponse({
            authenticated: false,
            user_id: targetUser || null,
            message: 'No authenticated user. Call pa-auth-start to begin Device Code Flow.',
          });
        }

        // Try to get a token to verify it's still valid
        const token = await userAuth.getAccessToken(user_id);
        if (token) {
          return jsonResponse({
            authenticated: true,
            user_id: targetUser,
            message: `Authenticated as ${targetUser}. Write operations are available.`,
          });
        }

        return jsonResponse({
          authenticated: false,
          user_id: targetUser,
          message: 'Token expired or invalid. Call pa-auth-start to re-authenticate.',
        });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  console.log(`[AuthTools] ${count} auth tools registered with McpServer (SSE transport)`);
  return count;
}
