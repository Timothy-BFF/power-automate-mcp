/**
 * Auth Tool Handlers — v3.0.0
 * 
 * MCP tool handler implementations for per-user delegated authentication.
 * These are called by the MCP server's tools/call dispatcher in index.ts.
 * 
 * Pattern mirrors Power Interpreter's ms_auth(action) tool.
 * 
 * @version 3.0.0
 */

import {
  UserAuthManager,
  AuthStartResult,
  AuthPollResult,
  AuthStatusResult,
} from '../auth/user-auth-manager';

// ——— Tool: pa-auth-start ———————————————————————————————————————

export async function handleAuthStart(
  args: { user_id: string },
  userAuthManager: UserAuthManager
): Promise<{ content: Array<{ type: string; text: string }> }> {
  try {
    const result: AuthStartResult = await userAuthManager.startAuth(args.user_id);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(
            {
              status: 'device_code_issued',
              user_id: args.user_id,
              user_code: result.userCode,
              verification_uri: result.verificationUri,
              expires_in_seconds: result.expiresIn,
              instructions: result.message,
              next_step:
                'After the user completes login at the URL above, call pa-auth-poll to finish authentication.',
            },
            null,
            2
          ),
        },
      ],
    };
  } catch (error: any) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(
            {
              status: 'error',
              error: error.message,
              suggestion: !userAuthManager.isConfigured()
                ? 'Delegated auth is not configured. Set PA_USER_CLIENT_ID and PA_USER_TENANT_ID in environment variables.'
                : 'Check Azure AD App Registration configuration and ensure Device Code Flow is enabled.',
            },
            null,
            2
          ),
        },
      ],
    };
  }
}

// ——— Tool: pa-auth-poll ————————————————————————————————————————

export async function handleAuthPoll(
  args: { user_id?: string },
  userAuthManager: UserAuthManager
): Promise<{ content: Array<{ type: string; text: string }> }> {
  try {
    const result: AuthPollResult = await userAuthManager.pollAuth(args.user_id);

    const response: Record<string, any> = {
      status: result.status,
      message: result.message,
    };

    if (result.userId) {
      response.user_id = result.userId;
    }

    // Add next-step guidance based on status
    switch (result.status) {
      case 'authenticated':
        response.next_step =
          'Authentication complete. Write operations (create/update/delete flows) are now available for this user.';
        break;
      case 'pending':
        response.next_step =
          'User has not completed login yet. Wait a few seconds and call pa-auth-poll again.';
        break;
      case 'expired':
        response.next_step =
          'Device code has expired. Call pa-auth-start to begin a new authentication.';
        break;
      case 'error':
        response.next_step =
          'An error occurred. Review the message and try pa-auth-start again if needed.';
        break;
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error: any) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(
            {
              status: 'error',
              error: error.message,
            },
            null,
            2
          ),
        },
      ],
    };
  }
}

// ——— Tool: pa-auth-status ——————————————————————————————————————

export function handleAuthStatus(
  args: { user_id?: string },
  userAuthManager: UserAuthManager
): { content: Array<{ type: string; text: string }> } {
  try {
    const result: AuthStatusResult = userAuthManager.getStatus(args.user_id);

    const response: Record<string, any> = {
      authenticated: result.authenticated,
      message: result.message,
    };

    if (result.userId) {
      response.user_id = result.userId;
    }

    if (result.expiresIn !== undefined) {
      response.expires_in_seconds = result.expiresIn;
      response.expires_in_minutes = Math.floor(result.expiresIn / 60);
    }

    // Include list of all authenticated users
    const authenticatedUsers = userAuthManager.listAuthenticatedUsers();
    if (authenticatedUsers.length > 0) {
      response.authenticated_users = authenticatedUsers;
    }

    // Include delegated auth configuration status
    response.delegated_auth_configured = userAuthManager.isConfigured();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error: any) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(
            {
              status: 'error',
              error: error.message,
            },
            null,
            2
          ),
        },
      ],
    };
  }
}

// ——— Tool Dispatcher ——————————————————————————————————————————

/**
 * Dispatch an auth tool call to the appropriate handler.
 * This function is designed to be called from the MCP server's
 * tools/call handler in index.ts.
 * 
 * @param toolName - One of 'pa-auth-start', 'pa-auth-poll', 'pa-auth-status'
 * @param args - The tool arguments from the MCP request
 * @param userAuthManager - The UserAuthManager instance
 * @returns MCP tool response
 */
export async function dispatchAuthTool(
  toolName: string,
  args: Record<string, any>,
  userAuthManager: UserAuthManager
): Promise<{ content: Array<{ type: string; text: string }> } | null> {
  switch (toolName) {
    case 'pa-auth-start':
      return handleAuthStart(args as { user_id: string }, userAuthManager);

    case 'pa-auth-poll':
      return handleAuthPoll(args as { user_id?: string }, userAuthManager);

    case 'pa-auth-status':
      return handleAuthStatus(args as { user_id?: string }, userAuthManager);

    default:
      return null; // Not an auth tool — let the existing handler process it
  }
}
