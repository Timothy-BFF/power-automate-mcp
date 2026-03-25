/**
 * Flow Adopter — Power Automate MCP v3.4.0
 *
 * Registers a flow in the Dataverse `workflows` entity so it can be
 * added to solutions via AddSolutionComponent.
 *
 * Background: Flows created via the Flow Management API or the portal UI
 * are "non-solution-aware" — they exist only in the Flow service storage,
 * not in Dataverse. The AddSolutionComponent action requires a Dataverse
 * `workflows` entity record to exist for the flow. This module creates
 * that record, effectively "adopting" the flow into the Dataverse world.
 *
 * This is the programmatic equivalent of the portal's
 * "Add existing → Automation → Cloud flow → Outside solutions" button.
 *
 * Directly addresses Jose's BeakSolution 404 error (2026-03-25):
 *   "Cannot add Workflow with id (b65abe28-...) to solution (...) because it does not exist"
 *   Root cause: Flow existed in Flow API but had no Dataverse workflows record.
 */

import axios from 'axios';

export interface RegisterFlowResult {
  registered: boolean;
  alreadyExisted: boolean;
  flowId: string;
  message: string;
}

/**
 * Registers a flow in the Dataverse `workflows` entity.
 *
 * Steps:
 * 1. Check if a workflow record already exists (GET /workflows(flowId))
 * 2. If not, create it (POST /workflows) with category=5 (Cloud Flow)
 * 3. Return status indicating whether the record was created or already existed
 *
 * After this function returns successfully, AddSolutionComponent will
 * be able to find the flow and add it to a solution.
 *
 * @param options - Flow metadata and Dataverse connection details
 * @returns RegisterFlowResult indicating success
 * @throws Error if Dataverse is unreachable or creation fails
 */
export async function registerFlowInDataverse(options: {
  flowId: string;
  displayName: string;
  dataverseUrl: string;
  dataverseToken: string;
  flowDefinition?: any;
  connectionReferences?: any;
}): Promise<RegisterFlowResult> {
  const { flowId, displayName, dataverseUrl, dataverseToken } = options;

  // Normalize URL — DATAVERSE_URL may or may not include https://
  const host = dataverseUrl.startsWith('https://') ? dataverseUrl : `https://${dataverseUrl}`;
  const baseUrl = `${host}/api/data/v9.2`;

  const headers: Record<string, string> = {
    'Authorization': `Bearer ${dataverseToken}`,
    'Content-Type': 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
  };

  // ─── Step 1: Check if workflow record already exists ───────────────────
  try {
    const check = await axios.get(`${baseUrl}/workflows(${flowId})`, {
      headers,
      params: { '$select': 'workflowid,name,category' },
    });
    console.log(`[FlowAdopter] Workflow record already exists in Dataverse: ${flowId} ("${check.data?.name}")`);
    return {
      registered: true,
      alreadyExisted: true,
      flowId,
      message: `Workflow record already exists in Dataverse ("${check.data?.name}"). Ready for AddSolutionComponent.`,
    };
  } catch (e: any) {
    if (e.response?.status !== 404) {
      const detail = e.response?.data?.error?.message || e.message;
      throw new Error(`[FlowAdopter] Failed to check Dataverse workflow record: ${e.response?.status || 'unknown'} — ${detail}`);
    }
    console.log(`[FlowAdopter] No Dataverse record for ${flowId} — creating now`);
  }

  // ─── Step 2: Create the workflow record ────────────────────────────────
  const workflowPayload: Record<string, any> = {
    workflowid: flowId,
    name: displayName,
    type: 1,           // 1 = Definition
    category: 5,       // 5 = Modern Flow (Cloud Flow)
    primaryentity: 'none',
  };

  // Include clientdata if we have the flow definition — this allows the
  // portal to display the flow's definition when opened from the solution.
  if (options.flowDefinition) {
    try {
      workflowPayload.clientdata = JSON.stringify({
        properties: {
          definition: options.flowDefinition,
          connectionReferences: options.connectionReferences || {},
        },
      });
      console.log(`[FlowAdopter] Including clientdata (${workflowPayload.clientdata.length} chars)`);
    } catch (jsonErr) {
      console.warn(`[FlowAdopter] Could not serialize clientdata — proceeding without it`);
    }
  }

  try {
    await axios.post(`${baseUrl}/workflows`, workflowPayload, { headers });
    console.log(`[FlowAdopter] Created Dataverse workflow record: ${flowId} ("${displayName}")`);
    return {
      registered: true,
      alreadyExisted: false,
      flowId,
      message: `Dataverse workflow record created for "${displayName}". Ready for AddSolutionComponent.`,
    };
  } catch (e: any) {
    const status = e.response?.status;
    const detail = e.response?.data?.error?.message || e.message;

    // Handle duplicate key — record was created between our check and insert (race condition)
    if (status === 409 || (detail && detail.toLowerCase().includes('duplicate'))) {
      console.log(`[FlowAdopter] Concurrent creation detected (${status}) — record exists: ${flowId}`);
      return {
        registered: true,
        alreadyExisted: true,
        flowId,
        message: `Workflow record created concurrently. Ready for AddSolutionComponent.`,
      };
    }

    throw new Error(`[FlowAdopter] Failed to create Dataverse workflow record (${status}): ${detail}`);
  }
}
