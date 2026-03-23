/**
 * SolutionClient — Dataverse Solutions API client
 *
 * Provides access to Dataverse Web API v9.2 for solution lifecycle management.
 * Uses service principal token (AzureTokenManager) with Dataverse scope.
 *
 * Operations:
 *   - List solutions (unmanaged or all)
 *   - Get solution details (by GUID or unique name)
 *   - List solution components
 *   - Export solution as base64 ZIP
 *   - Add component to solution (e.g., add a Cloud Flow)
 *
 * Prerequisites:
 *   - DATAVERSE_URL env var (e.g., org99066845.crm.dynamics.com)
 *   - Service principal registered as Application User in Dataverse
 *   - System Customizer role assigned to the Application User
 *
 * @version 3.2.0
 */

import axios from 'axios';
import { AzureTokenManager } from '../auth/azure-token-manager.js';

// =============================================================================
// Constants
// =============================================================================

const DATAVERSE_API_VERSION = 'v9.2';

/** Dataverse component type codes → human-readable names */
export const COMPONENT_TYPES: Record<number, string> = {
  1: 'Entity',
  2: 'Attribute',
  3: 'Relationship',
  9: 'OptionSet',
  10: 'EntityRelationship',
  26: 'View',
  29: 'Workflow',
  60: 'SystemForm',
  61: 'WebResource',
  62: 'SiteMap',
  63: 'ConnectionRole',
  65: 'HierarchyRule',
  66: 'CustomControl',
  80: 'ModelDrivenApp',
  300: 'CanvasApp',
  371: 'Connector',
  372: 'EnvironmentVariableDefinition',
  373: 'EnvironmentVariableValue',
  380: 'AIModel',
  381: 'AITemplate',
  400: 'DesktopFlow',
};

// =============================================================================
// Exported Types
// =============================================================================

export interface SolutionSummary {
  solutionId: string;
  uniqueName: string;
  friendlyName: string;
  version: string;
  isManaged: boolean;
  installedOn: string | null;
  modifiedOn: string | null;
  description: string | null;
  publisherId: string | null;
}

export interface SolutionComponent {
  componentId: string;
  objectId: string;
  componentType: number;
  componentTypeName: string;
  isMetadata: boolean;
}

export interface ExportResult {
  fileName: string;
  base64Content: string;
  sizeBytes: number;
}

// =============================================================================
// SolutionClient
// =============================================================================

export class SolutionClient {
  private tokenManager: AzureTokenManager;
  private dataverseUrl: string;
  private baseUrl: string;
  private scope: string;

  constructor(tokenManager: AzureTokenManager) {
    this.tokenManager = tokenManager;

    const envUrl = (process.env.DATAVERSE_URL || '').trim();
    if (!envUrl) {
      console.warn('[SolutionClient] DATAVERSE_URL not set — Dataverse operations disabled');
      this.dataverseUrl = '';
      this.baseUrl = '';
      this.scope = '';
    } else {
      // Normalize: ensure https:// prefix and no trailing slash
      this.dataverseUrl = envUrl.startsWith('https://') ? envUrl : `https://${envUrl}`;
      this.dataverseUrl = this.dataverseUrl.replace(/\/+$/, '');
      this.baseUrl = `${this.dataverseUrl}/api/data/${DATAVERSE_API_VERSION}`;
      this.scope = `${this.dataverseUrl}/.default`;
      console.log(`[SolutionClient] Dataverse configured: ${this.dataverseUrl}`);
      console.log(`[SolutionClient]   Scope: ${this.scope}`);
    }
  }

  // -----------------------------------------------------------------------
  // Configuration
  // -----------------------------------------------------------------------

  /** Returns true if DATAVERSE_URL is configured. */
  isConfigured(): boolean {
    return !!this.dataverseUrl;
  }

  /** Returns the Dataverse token scope for pre-warming. */
  getScope(): string {
    return this.scope;
  }

  // -----------------------------------------------------------------------
  // Internal Request Helper
  // -----------------------------------------------------------------------

  private async dataverseRequest(
    path: string,
    method: string = 'GET',
    body?: any
  ): Promise<any> {
    if (!this.isConfigured()) {
      throw new Error(
        'Dataverse not configured. Set DATAVERSE_URL environment variable ' +
        '(e.g., org99066845.crm.dynamics.com).'
      );
    }

    const token = await this.tokenManager.getToken(this.scope);
    const url = `${this.baseUrl}${path}`;

    console.log(`[SolutionClient] ${method} ${path}`);

    const response = await axios({
      method,
      url,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      data: body,
      validateStatus: (s) => s < 500,
    });

    if (response.status >= 400) {
      const errMsg =
        response.data?.error?.message ||
        response.data?.Message ||
        `HTTP ${response.status}`;
      throw new Error(`Dataverse API error (${response.status}): ${errMsg}`);
    }

    return response.data;
  }

  // -----------------------------------------------------------------------
  // Solution Operations
  // -----------------------------------------------------------------------

  /**
   * Lists all solutions in the Dataverse environment.
   * By default shows only unmanaged, visible solutions.
   */
  async listSolutions(includeManaged: boolean = false): Promise<SolutionSummary[]> {
    let filter = 'isvisible eq true';
    if (!includeManaged) {
      filter += ' and ismanaged eq false';
    }

    const path =
      `/solutions?$select=solutionid,uniquename,friendlyname,version,ismanaged,` +
      `installedon,modifiedon,description,_publisherid_value` +
      `&$filter=${encodeURIComponent(filter)}` +
      `&$orderby=friendlyname asc`;

    const data = await this.dataverseRequest(path);
    const solutions = data.value || [];

    console.log(`[SolutionClient] Found ${solutions.length} solutions`);

    return solutions.map((s: any) => ({
      solutionId: s.solutionid,
      uniqueName: s.uniquename,
      friendlyName: s.friendlyname,
      version: s.version,
      isManaged: s.ismanaged,
      installedOn: s.installedon || null,
      modifiedOn: s.modifiedon || null,
      description: s.description || null,
      publisherId: s._publisherid_value || null,
    }));
  }

  /**
   * Gets details for a specific solution.
   * Accepts either a GUID (solutionid) or a unique name.
   */
  async getSolution(solutionIdOrName: string): Promise<SolutionSummary> {
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
      solutionIdOrName
    );

    let path: string;
    if (isGuid) {
      path =
        `/solutions(${solutionIdOrName})?$select=solutionid,uniquename,friendlyname,` +
        `version,description,ismanaged,installedon,modifiedon,_publisherid_value`;
    } else {
      path =
        `/solutions?$select=solutionid,uniquename,friendlyname,version,description,` +
        `ismanaged,installedon,modifiedon,_publisherid_value` +
        `&$filter=uniquename eq '${solutionIdOrName}'`;
    }

    const data = await this.dataverseRequest(path);

    // If queried by unique name, extract from value array
    const s = isGuid ? data : data.value?.[0] || null;
    if (!s) {
      throw new Error(`Solution not found: ${solutionIdOrName}`);
    }

    return {
      solutionId: s.solutionid,
      uniqueName: s.uniquename,
      friendlyName: s.friendlyname,
      version: s.version,
      isManaged: s.ismanaged,
      installedOn: s.installedon || null,
      modifiedOn: s.modifiedon || null,
      description: s.description || null,
      publisherId: s._publisherid_value || null,
    };
  }

  /**
   * Lists components within a solution.
   * Returns component type, object ID, and human-readable type name.
   */
  async listSolutionComponents(solutionId: string): Promise<SolutionComponent[]> {
    const path =
      `/solutioncomponents?$filter=_solutionid_value eq '${solutionId}'` +
      `&$select=solutioncomponentid,componenttype,objectid,ismetadata` +
      `&$orderby=componenttype asc`;

    const data = await this.dataverseRequest(path);
    const components = data.value || [];

    console.log(`[SolutionClient] Solution ${solutionId}: ${components.length} components`);

    return components.map((c: any) => ({
      componentId: c.solutioncomponentid,
      objectId: c.objectid,
      componentType: c.componenttype,
      componentTypeName: COMPONENT_TYPES[c.componenttype] || `Unknown(${c.componenttype})`,
      isMetadata: c.ismetadata || false,
    }));
  }

  /**
   * Exports a solution as a base64-encoded ZIP file.
   * Uses the Dataverse ExportSolution action.
   */
  async exportSolution(
    solutionUniqueName: string,
    managed: boolean = false
  ): Promise<ExportResult> {
    console.log(`[SolutionClient] Exporting solution: ${solutionUniqueName} (managed: ${managed})`);

    const body = {
      SolutionName: solutionUniqueName,
      Managed: managed,
    };

    const data = await this.dataverseRequest('/ExportSolution', 'POST', body);

    const base64 = data.ExportSolutionFile || '';
    const sizeBytes = Math.round((base64.length * 3) / 4);
    const suffix = managed ? '_managed' : '';

    console.log(`[SolutionClient] Exported ${solutionUniqueName}: ${Math.round(sizeBytes / 1024)}KB`);

    return {
      fileName: `${solutionUniqueName}${suffix}.zip`,
      base64Content: base64,
      sizeBytes,
    };
  }

  /**
   * Adds a component (e.g., a Cloud Flow) to a solution.
   * Uses the Dataverse AddSolutionComponent action.
   *
   * Common component types:
   *   29 = Workflow / Cloud Flow
   *   1  = Entity / Table
   *   300 = Canvas App
   *   372 = Environment Variable Definition
   */
  async addSolutionComponent(
    solutionUniqueName: string,
    componentId: string,
    componentType: number = 29,
    addRequiredComponents: boolean = false
  ): Promise<{ success: boolean; message: string }> {
    const typeName = COMPONENT_TYPES[componentType] || `Type(${componentType})`;

    console.log(
      `[SolutionClient] Adding ${typeName} ${componentId} to solution ${solutionUniqueName}`
    );

    const body = {
      ComponentId: componentId,
      ComponentType: componentType,
      SolutionUniqueName: solutionUniqueName,
      AddRequiredComponents: addRequiredComponents,
    };

    await this.dataverseRequest('/AddSolutionComponent', 'POST', body);

    console.log(`[SolutionClient] ✅ ${typeName} added to ${solutionUniqueName}`);

    return {
      success: true,
      message: `${typeName} ${componentId} added to solution ${solutionUniqueName}`,
    };
  }
}
