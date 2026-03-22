/**
 * Power Automate MCP - Auth Tool Handlers (v3.0.3)
 *
 * Registers the 3 Device Code Flow authentication tools on McpServer (SSE).
 * Descriptions imported from tool-descriptions.ts.
 *
 * UserAuthManager method mapping:
 *   pa-auth-start  -> userAuth.startAuth(userId)
 *   pa-auth-poll   -> userAuth.pollAuth(userId)
 *   pa-auth-status -> userAuth.getAuthStatus(userId?)
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
// Auth Tool Registration (SSE transport via McpServer)
// =========================================================================

export function registerAuthTools(
  server: McpServer,
  userAuth: UserAuthManager
): number {
  let count = 0;

  // -----------------------------------------------------------------------
  // pa-auth-start: Initiate Device Code Flow
  // UserAuthManager.startAuth(userId) -> AuthStartResult
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-start',
    TOOL_DESCRIPTIONS['pa-auth-start'],
    {
      user_id: z.string().describe('User email (e.g., jose@company.com). Required.'),
    },
    async ({ user_id }) => {
      try {
        if (!user_id) {
          return errorResponse({ message: 'user_id is required. Provide the user email address.' });
        }
        const result = await userAuth.startAuth(user_id);
        return jsonResponse({
          status: 'device_code_issued',
          user_code: result.userCode,
          verification_uri: result.verificationUri,
          message: result.message,
          expires_in: result.expiresIn,
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
  // UserAuthManager.pollAuth(userId) -> AuthPollResult
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-poll',
    TOOL_DESCRIPTIONS['pa-auth-poll'],
    {
      user_id: z.string().optional().describe('User email to poll for. If omitted, polls for the most recent pending auth.'),
    },
    async ({ user_id }) => {
      try {
        const targetUser = user_id || userAuth.getDefaultUserId();
        if (!targetUser) {
          return jsonResponse({
            status: 'error',
            message: 'No user_id provided and no pending authentication found. Call pa-auth-start first.',
          });
        }

        const result = await userAuth.pollAuth(targetUser);

        if (result.status === 'authenticated') {
          return jsonResponse({
            status: 'authenticated',
            user_id: result.userId || targetUser,
            message: result.message,
            _note: 'You can now use pa-create-flow, pa-update-flow, and pa-trigger-flow.',
          });
        }

        if (result.status === 'pending') {
          return jsonResponse({
            status: 'pending',
            message: result.message,
          });
        }

        // expired or error
        return jsonResponse({
          status: result.status,
          message: result.message,
        });
      } catch (error: any) {
        return errorResponse(error);
      }
    }
  );
  count++;

  // -----------------------------------------------------------------------
  // pa-auth-status: Check current authentication state
  // UserAuthManager.getAuthStatus(userId?) -> AuthStatusResult
  // -----------------------------------------------------------------------
  server.tool(
    'pa-auth-status',
    TOOL_DESCRIPTIONS['pa-auth-status'],
    {
      user_id: z.string().optional().describe('User email to check. If omitted, returns general status.'),
    },
    async ({ user_id }) => {
      try {
        const status = userAuth.getAuthStatus(user_id);
        return jsonResponse({
          authenticated: status.authenticated,
          user_id: status.userId,
          message: status.message,
          pending: status.pending,
          authenticatedUsers: status.authenticatedUsers,
          expiresIn: status.expiresIn,
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
