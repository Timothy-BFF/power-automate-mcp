/**
 * Tool Descriptions — Power Automate MCP
 * v3.3.0: 23 MCP tools (+ 3 auth tools registered separately)
 *
 * v3.2.0 → v3.3.0 changes:
 *   FIX: pa-list-connections (endpoint path corrected)
 *   NEW: pa-create-solution, pa-delete-solution, pa-create-connection
 */

export interface ToolDescription {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, any>;
    required?: string[];
  };
}

export const TOOL_DESCRIPTIONS: ToolDescription[] = [
  // ═══════════════════════════════════════════
  // FLOW MANAGEMENT (7 tools)
  // ═══════════════════════════════════════════
  {
    name: 'pa-list-flows',
    description: 'List all flows in a Power Platform environment. Returns flow names, IDs, states, and trigger information.',
    inputSchema: {
      type: 'object',
      properties: {
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty to use the default environment.',
        },
        filter: {
          type: 'string',
          description: 'Optional filter: "team" for shared/team flows, "personal" for personal flows.',
        },
      },
      required: [],
    },
  },
  {
    name: 'pa-get-flow-details',
    description: 'Get detailed information about a specific flow, including its full definition, triggers, actions, and connections.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to retrieve details for.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of authenticated user (optional — enables delegated token for full definition).',
        },
      },
      required: ['flow_id'],
    },
  },
  {
    name: 'pa-create-flow',
    description: 'Create a new Power Automate flow. Requires user authentication (pa-auth-start). Provide the FULL flow definition in one call.',
    inputSchema: {
      type: 'object',
      properties: {
        display_name: {
          type: 'string',
          description: 'Display name for the new flow.',
        },
        definition: {
          type: 'object',
          description: 'Full flow definition JSON including triggers and actions.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user who will own this flow.',
        },
      },
      required: ['display_name', 'definition', 'user_id'],
    },
  },
  {
    name: 'pa-update-flow',
    description: 'Update an existing Power Automate flow. Can update display name, definition, or both.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to update.',
        },
        display_name: {
          type: 'string',
          description: 'New display name (optional).',
        },
        definition: {
          type: 'object',
          description: 'Updated flow definition JSON (optional).',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user.',
        },
      },
      required: ['flow_id', 'user_id'],
    },
  },
  {
    name: 'pa-delete-flow',
    description: 'Delete a Power Automate flow permanently.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to delete.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user.',
        },
      },
      required: ['flow_id'],
    },
  },
  {
    name: 'pa-enable-flow',
    description: 'Enable (turn on) a Power Automate flow.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to enable.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
      },
      required: ['flow_id'],
    },
  },
  {
    name: 'pa-disable-flow',
    description: 'Disable (turn off) a Power Automate flow.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to disable.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
      },
      required: ['flow_id'],
    },
  },

  // ═══════════════════════════════════════════
  // FLOW RUN MANAGEMENT (3 tools)
  // ═══════════════════════════════════════════
  {
    name: 'pa-get-flow-runs',
    description: 'Get the run history for a specific flow. Returns recent runs with status, start/end times, and trigger info.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID) to get runs for.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        top: {
          type: 'number',
          description: 'Number of runs to return (default: 25, max: 50).',
        },
      },
      required: ['flow_id'],
    },
  },
  {
    name: 'pa-get-flow-run-details',
    description: 'Get detailed information about a specific flow run, including action results and error details.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID).',
        },
        run_id: {
          type: 'string',
          description: 'The run ID (GUID) to get details for.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
      },
      required: ['flow_id', 'run_id'],
    },
  },
  {
    name: 'pa-resubmit-flow-run',
    description: 'Resubmit (retry) a failed or cancelled flow run using the original trigger data.',
    inputSchema: {
      type: 'object',
      properties: {
        flow_id: {
          type: 'string',
          description: 'The flow ID (GUID).',
        },
        run_id: {
          type: 'string',
          description: 'The run ID (GUID) to resubmit.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user.',
        },
      },
      required: ['flow_id', 'run_id'],
    },
  },

  // ═══════════════════════════════════════════
  // ENVIRONMENT MANAGEMENT (2 tools)
  // ═══════════════════════════════════════════
  {
    name: 'pa-list-environments',
    description: 'List all Power Platform environments accessible to the service principal.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  {
    name: 'pa-get-environment',
    description: 'Get detailed information about a specific Power Platform environment.',
    inputSchema: {
      type: 'object',
      properties: {
        environment_id: {
          type: 'string',
          description: 'The environment ID to get details for.',
        },
      },
      required: ['environment_id'],
    },
  },

  // ═══════════════════════════════════════════
  // CONNECTION MANAGEMENT (4 tools — 3 existing + 1 new)
  // ═══════════════════════════════════════════
  {
    name: 'pa-list-connections',
    description: 'List all connections in a Power Platform environment. v3.3.0: Fixed to use delegated endpoint path.',
    inputSchema: {
      type: 'object',
      properties: {
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user (optional — uses delegated token for better results).',
        },
      },
      required: [],
    },
  },
  {
    name: 'pa-get-connection',
    description: 'Get detailed information about a specific connection.',
    inputSchema: {
      type: 'object',
      properties: {
        connection_id: {
          type: 'string',
          description: 'The connection ID (name) to retrieve.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user (optional).',
        },
      },
      required: ['connection_id'],
    },
  },
  {
    name: 'pa-list-connectors',
    description: 'List available connectors (APIs) in a Power Platform environment.',
    inputSchema: {
      type: 'object',
      properties: {
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
      },
      required: [],
    },
  },
  {
    name: 'pa-create-connection',
    description: 'Create a new connection to a connector in Power Automate. Requires user authentication (pa-auth-start). The connection will be owned by the authenticated user.',
    inputSchema: {
      type: 'object',
      properties: {
        connector_id: {
          type: 'string',
          description: 'The connector API name (e.g., "shared_office365", "shared_sharepointonline") or full API ID path.',
        },
        environment_id: {
          type: 'string',
          description: 'Power Platform environment ID. Leave empty for default.',
        },
        user_id: {
          type: 'string',
          description: 'Email of the authenticated user who will own this connection.',
        },
        connection_parameters: {
          type: 'object',
          description: 'Optional connection parameters (varies by connector). Most OAuth connectors need no extra params.',
        },
      },
      required: ['connector_id', 'user_id'],
    },
  },

  // ═══════════════════════════════════════════
  // SOLUTION MANAGEMENT (7 tools — 5 existing + 2 new)
  // ═══════════════════════════════════════════
  {
    name: 'pa-list-solutions',
    description: 'List Dataverse solutions. By default shows only unmanaged solutions.',
    inputSchema: {
      type: 'object',
      properties: {
        unmanaged_only: {
          type: 'boolean',
          description: 'If true (default), only return unmanaged solutions. Set false to include managed.',
        },
      },
      required: [],
    },
  },
  {
    name: 'pa-get-solution',
    description: 'Get detailed information about a specific Dataverse solution by unique name or GUID.',
    inputSchema: {
      type: 'object',
      properties: {
        unique_name: {
          type: 'string',
          description: 'Solution unique name (e.g., "MySolution") or solution GUID.',
        },
      },
      required: ['unique_name'],
    },
  },
  {
    name: 'pa-list-solution-components',
    description: 'List all components (flows, entities, web resources, etc.) within a Dataverse solution.',
    inputSchema: {
      type: 'object',
      properties: {
        solution_id: {
          type: 'string',
          description: 'The solution GUID (solutionId from pa-get-solution).',
        },
      },
      required: ['solution_id'],
    },
  },
  {
    name: 'pa-add-solution-component',
    description: 'Add an existing component (flow, entity, etc.) to a Dataverse solution.',
    inputSchema: {
      type: 'object',
      properties: {
        solution_unique_name: {
          type: 'string',
          description: 'The unique name of the target solution.',
        },
        component_id: {
          type: 'string',
          description: 'GUID of the component to add.',
        },
        component_type: {
          type: 'number',
          description: 'Component type code (e.g., 29 = Workflow, 1 = Entity, 300 = Canvas App, 62 = WebResource).',
        },
        add_required_components: {
          type: 'boolean',
          description: 'If true, also add required dependent components. Default: false.',
        },
      },
      required: ['solution_unique_name', 'component_id', 'component_type'],
    },
  },
  {
    name: 'pa-export-solution',
    description: 'Export a Dataverse solution as a ZIP file (base64 encoded).',
    inputSchema: {
      type: 'object',
      properties: {
        solution_name: {
          type: 'string',
          description: 'The unique name of the solution to export.',
        },
        managed: {
          type: 'boolean',
          description: 'Export as managed (true) or unmanaged (false, default).',
        },
      },
      required: ['solution_name'],
    },
  },
  {
    name: 'pa-create-solution',
    description: 'Create a new Dataverse solution. Requires a publisher ID (use pa-list-solutions to find existing publishers).',
    inputSchema: {
      type: 'object',
      properties: {
        unique_name: {
          type: 'string',
          description: 'Unique name for the solution (no spaces, e.g., "MyNewSolution").',
        },
        friendly_name: {
          type: 'string',
          description: 'Display name for the solution (e.g., "My New Solution").',
        },
        publisher_id: {
          type: 'string',
          description: 'GUID of the publisher. Use publisherId from pa-get-solution on an existing solution to find valid publishers.',
        },
        version: {
          type: 'string',
          description: 'Version number (default: "1.0.0.0").',
        },
        description: {
          type: 'string',
          description: 'Optional description of the solution.',
        },
      },
      required: ['unique_name', 'friendly_name', 'publisher_id'],
    },
  },
  {
    name: 'pa-delete-solution',
    description: 'Delete a Dataverse solution. WARNING: This permanently removes the solution container. Components inside may remain in the environment.',
    inputSchema: {
      type: 'object',
      properties: {
        solution_id: {
          type: 'string',
          description: 'The solution GUID to delete (solutionId from pa-get-solution).',
        },
      },
      required: ['solution_id'],
    },
  },
];
