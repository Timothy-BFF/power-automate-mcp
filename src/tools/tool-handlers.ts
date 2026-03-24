/**
 * Tool Handlers — Power Automate MCP
 * v3.3.0: 23 MCP tool handlers
 *
 * v3.2.0 → v3.3.0 changes:
 *   FIX: pa-list-connections now routes through fixed ConnectionClient
 *   NEW: pa-create-solution, pa-delete-solution, pa-create-connection
 */

export interface ToolClients {
  flowClient: any;
  connectionClient: any;
  environmentClient: any;
  solutionClient: any;
}

function resolveEnv(explicit?: string): string {
  if (explicit) {
    console.log(`[EnvResolver] Using explicit parameter: ${explicit}`);
    return explicit;
  }
  const envId = process.env.POWER_PLATFORM_ENVIRONMENT_ID || '';
  console.log(`[EnvResolver] Using POWER_PLATFORM_ENVIRONMENT_ID: ${envId}`);
  return envId;
}

function success(data: any): { content: Array<{ type: string; text: string }> } {
  return {
    content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
  };
}

function error(message: string): { content: Array<{ type: string; text: string }>; isError: boolean } {
  return {
    content: [{ type: 'text', text: JSON.stringify({ error: message }, null, 2) }],
    isError: true,
  };
}

export async function handleToolCall(
  toolName: string,
  args: Record<string, unknown>,
  clients: ToolClients
): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
  try {
    const envId = resolveEnv(args.environment_id as string | undefined);

    switch (toolName) {
      // ═══════════════════════════════════════════
      // FLOW MANAGEMENT
      // ═══════════════════════════════════════════
      case 'pa-list-flows': {
        const result = await clients.flowClient.listFlows(envId, args.filter as string | undefined);
        return success(result);
      }

      case 'pa-get-flow-details': {
        const result = await clients.flowClient.getFlowDetails(
          envId,
          args.flow_id as string,
          args.user_id as string | undefined
        );
        return success(result);
      }

      case 'pa-create-flow': {
        const result = await clients.flowClient.createFlow(
          envId,
          args.display_name as string,
          args.definition,
          args.user_id as string
        );
        return success(result);
      }

      case 'pa-update-flow': {
        const result = await clients.flowClient.updateFlow(
          envId,
          args.flow_id as string,
          {
            displayName: args.display_name as string | undefined,
            definition: args.definition,
          },
          args.user_id as string | undefined
        );
        return success(result);
      }

      case 'pa-delete-flow': {
        const result = await clients.flowClient.deleteFlow(
          envId,
          args.flow_id as string,
          args.user_id as string | undefined
        );
        return success(result);
      }

      case 'pa-enable-flow': {
        const result = await clients.flowClient.enableFlow(
          envId,
          args.flow_id as string
        );
        return success(result);
      }

      case 'pa-disable-flow': {
        const result = await clients.flowClient.disableFlow(
          envId,
          args.flow_id as string
        );
        return success(result);
      }

      // ═══════════════════════════════════════════
      // FLOW RUN MANAGEMENT
      // ═══════════════════════════════════════════
      case 'pa-get-flow-runs': {
        const result = await clients.flowClient.getFlowRuns(
          envId,
          args.flow_id as string,
          { top: args.top as number | undefined }
        );
        return success(result);
      }

      case 'pa-get-flow-run-details': {
        const result = await clients.flowClient.getFlowRunDetails(
          envId,
          args.flow_id as string,
          args.run_id as string
        );
        return success(result);
      }

      case 'pa-resubmit-flow-run': {
        const result = await clients.flowClient.resubmitFlowRun(
          envId,
          args.flow_id as string,
          args.run_id as string,
          args.user_id as string | undefined
        );
        return success(result);
      }

      // ═══════════════════════════════════════════
      // ENVIRONMENT MANAGEMENT
      // ═══════════════════════════════════════════
      case 'pa-list-environments': {
        const result = await clients.environmentClient.listEnvironments();
        return success(result);
      }

      case 'pa-get-environment': {
        const result = await clients.environmentClient.getEnvironment(
          args.environment_id as string
        );
        return success(result);
      }

      // ═══════════════════════════════════════════
      // CONNECTION MANAGEMENT
      // ═══════════════════════════════════════════
      case 'pa-list-connections': {
        // v3.3.0: Now routes through fixed ConnectionClient (delegated path)
        const result = await clients.connectionClient.listConnections(
          envId,
          args.user_id as string | undefined
        );
        return success(result);
      }

      case 'pa-get-connection': {
        const result = await clients.connectionClient.getConnection(
          envId,
          args.connection_id as string,
          args.user_id as string | undefined
        );
        return success(result);
      }

      case 'pa-list-connectors': {
        const result = await clients.connectionClient.listConnectors(envId);
        return success(result);
      }

      case 'pa-create-connection': {
        // v3.3.0: NEW — requires user delegated token
        const result = await clients.connectionClient.createConnection(
          envId,
          args.connector_id as string,
          args.user_id as string,
          args.connection_parameters as Record<string, any> | undefined
        );
        return success(result);
      }

      // ═══════════════════════════════════════════
      // SOLUTION MANAGEMENT (Dataverse)
      // ═══════════════════════════════════════════
      case 'pa-list-solutions': {
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const unmanagedOnly = args.unmanaged_only !== false; // default true
        const result = await clients.solutionClient.listSolutions(unmanagedOnly);
        return success(result);
      }

      case 'pa-get-solution': {
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.getSolution(
          args.unique_name as string
        );
        return success(result);
      }

      case 'pa-list-solution-components': {
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.listSolutionComponents(
          args.solution_id as string
        );
        return success(result);
      }

      case 'pa-add-solution-component': {
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.addSolutionComponent(
          args.solution_unique_name as string,
          args.component_id as string,
          args.component_type as number,
          args.add_required_components as boolean | undefined
        );
        return success(result);
      }

      case 'pa-export-solution': {
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.exportSolution(
          args.solution_name as string,
          args.managed as boolean | undefined
        );
        return success(result);
      }

      case 'pa-create-solution': {
        // v3.3.0: NEW
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.createSolution(
          args.unique_name as string,
          args.friendly_name as string,
          args.publisher_id as string,
          (args.version as string) || '1.0.0.0',
          (args.description as string) || ''
        );
        return success(result);
      }

      case 'pa-delete-solution': {
        // v3.3.0: NEW
        if (!clients.solutionClient?.configured) {
          return error('Dataverse not configured. Set DATAVERSE_URL environment variable.');
        }
        const result = await clients.solutionClient.deleteSolution(
          args.solution_id as string
        );
        return success(result);
      }

      // ═══════════════════════════════════════════
      // UNKNOWN TOOL
      // ═══════════════════════════════════════════
      default:
        return error(`Unknown tool: ${toolName}. Available tools: 23 MCP + 3 auth.`);
    }
  } catch (err: any) {
    console.error(`[ToolHandler] ${toolName} failed:`, err.message);
    return error(`${toolName} failed: ${err.message}`);
  }
}
