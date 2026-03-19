/**
 * MCP Tool Definitions for User Authentication (v3.0.0)
 * 
 * These 3 tools mirror the Power Interpreter's ms_auth(action) pattern,
 * adapted for the Power Automate scope.
 */

export const AUTH_TOOL_DEFINITIONS = [
  {
    name: 'pa-auth-start',
    description:
      'Start user authentication for Power Automate write operations. ' +
      'Initiates a Device Code Flow — returns a user code and URL. ' +
      'The user must visit the URL and enter the code to complete login. ' +
      'Required before create/update/delete flow operations.',
    inputSchema: {
      type: 'object' as const,
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
  },
  {
    name: 'pa-auth-poll',
    description:
      'Poll for completion of a pending device code authentication. ' +
      'Call this after pa-auth-start once the user has entered their code. ' +
      'Returns "authenticated" on success, "pending" if still waiting, or "expired" if timed out.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        user_id: {
          type: 'string',
          description:
            'Email of the user being authenticated. ' +
            'Optional if only one auth is pending.',
        },
      },
      required: [],
    },
  },
  {
    name: 'pa-auth-status',
    description:
      'Check the current authentication status for Power Automate user operations. ' +
      'Shows whether a user has a valid delegated token, when it expires, ' +
      'and lists all authenticated users if no user_id is specified.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        user_id: {
          type: 'string',
          description:
            'Email of the user to check status for. ' +
            'If omitted, returns status for the most recent or all authenticated users.',
        },
      },
      required: [],
    },
  },
];
