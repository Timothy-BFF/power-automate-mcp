/**
 * Solution Client — Power Automate MCP (Dataverse Web API)
 * v3.2.0: Initial (list, get, components, add-component, export)
 * v3.3.0: Added createSolution, deleteSolution, fixed component type 10089
 *
 * Constructor takes (tokenManager) only — reads DATAVERSE_URL from env.
 * Exposes isConfigured() and getScope() for index.ts compatibility.
 */
import axios from 'axios';

interface TokenProvider {
  getToken(scope: string): Promise<string>;
}

// Dataverse solution component type mapping
const COMPONENT_TYPES: Record<number, string> = {
  1: 'Entity',
  2: 'Attribute',
  3: 'Relationship',
  9: 'OptionSet',
  10: 'EntityRelationship',
  20: 'Role',
  24: 'FieldPermission',
  25: 'FieldSecurityProfile',
  26: 'EntityMap',
  29: 'Workflow',
  31: 'Report',
  36: 'MailMergeTemplate',
  44: 'DuplicateDetectionRule',
  59: 'SavedQuery',
  60: 'SavedQueryVisualization',
  61: 'SystemForm',
  62: 'WebResource',
  63: 'SiteMap',
  65: 'ConnectionRole',
  70: 'ManagedProperty',
  90: 'PluginType',
  91: 'PluginAssembly',
  92: 'SDKMessageProcessingStep',
  95: 'ServiceEndpoint',
  150: 'RoutingRule',
  152: 'SLA',
  154: 'ConvertRule',
  161: 'MobileOfflineProfile',
  300: 'CanvasApp',
  371: 'Connector',
  372: 'EnvironmentVariableDefinition',
  373: 'EnvironmentVariableValue',
  380: 'AIProject',
  381: 'AIModel',
  10003: 'Team',
  10029: 'ConnectionReference',
  10076: 'DesktopFlow',
  10089: 'ModernFlow',
};

function componentTypeName(type: number): string {
  return COMPONENT_TYPES[type] || `Unknown(${type})`;
}

export class SolutionClient {
  private dataverseUrl: string;
  private scope: string;
  private apiBase: string;
  private tokenProvider: TokenProvider;
  private _configured: boolean;

  /**
   * Constructor — takes tokenManager only (1 argument).
   * Reads DATAVERSE_URL from process.env internally.
   * This matches the call in index.ts: new SolutionClient(tokenManager)
   */
  constructor(tokenProvider: TokenProvider) {
    const dataverseUrl = process.env.DATAVERSE_URL;
    this.tokenProvider = tokenProvider;

    if (!dataverseUrl) {
      this._configured = false;
      this.dataverseUrl = '';
      this.scope = '';
      this.apiBase = '';
      console.log('[SolutionClient] Dataverse not configured (DATAVERSE_URL missing)');
      return;
    }

    this.dataverseUrl = dataverseUrl.startsWith('https://') ? dataverseUrl : `https://${dataverseUrl}`;
    this.scope = `${this.dataverseUrl}/.default`;
    this.apiBase = `${this.dataverseUrl}/api/data/v9.2`;
    this._configured = true;

    console.log(`[SolutionClient] Dataverse configured: ${this.dataverseUrl}`);
    console.log(`[SolutionClient]   Scope: ${this.scope}`);
  }

  // -----------------------------------------------------------------------
  // Config accessors (called by index.ts)
  // -----------------------------------------------------------------------

  /**
   * Returns true if DATAVERSE_URL is set.
   * Called by index.ts: solutionClient.isConfigured()
   */
  isConfigured(): boolean {
    return this._configured;
  }

  /**
   * Returns the Dataverse token scope.
   * Called by index.ts: solutionClient.getScope()
   */
  getScope(): string {
    return this.scope;
  }

  // -----------------------------------------------------------------------
  // Internal helpers
  // -----------------------------------------------------------------------

  private async getToken(): Promise<string> {
    if (!this._configured) throw new Error('Dataverse not configured. Set DATAVERSE_URL.');
    return this.tokenProvider.getToken(this.scope);
  }

  private async headers(): Promise<Record<string, string>> {
    const token = await this.getToken();
    return {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Accept: 'application/json',
      Prefer: 'odata.include-annotations="*"',
    };
  }

  // -----------------------------------------------------------------------
  // List Solutions (returns array — matches index.ts handler)
  // -----------------------------------------------------------------------

  async listSolutions(includeManaged: boolean = false): Promise<any[]> {
    const select = 'solutionid,uniquename,friendlyname,version,ismanaged,installedon,modifiedon,description,_publisherid_value';
    let filter = 'isvisible eq true';
    if (!includeManaged) filter += ' and ismanaged eq false';

    const url = `${this.apiBase}/solutions?$select=${select}&$filter=${encodeURIComponent(filter)}&$orderby=friendlyname asc`;
    console.log(`[SolutionClient] GET /solutions?$select=${select}&$filter=${filter}&$orderby=friendlyname asc`);

    const response = await axios.get(url, { headers: await this.headers() });
    const solutions = response.data?.value || [];
    console.log(`[SolutionClient] Found ${solutions.length} solutions`);

    return solutions.map((s: any) => ({
      solutionId: s.solutionid,
      uniqueName: s.uniquename,
      friendlyName: s.friendlyname,
      version: s.version,
      isManaged: s.ismanaged,
      description: s.description || '',
      publisherId: s._publisherid_value,
      installedOn: s.installedon,
      modifiedOn: s.modifiedon,
    }));
  }

  // -----------------------------------------------------------------------
  // Get Solution (returns single object)
  // -----------------------------------------------------------------------

  async getSolution(uniqueNameOrId: string): Promise<any> {
    const select = 'solutionid,uniquename,friendlyname,version,description,ismanaged,installedon,modifiedon,_publisherid_value';
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(uniqueNameOrId);

    let url: string;
    if (isGuid) {
      url = `${this.apiBase}/solutions(${uniqueNameOrId})?$select=${select}`;
    } else {
      url = `${this.apiBase}/solutions?$select=${select}&$filter=${encodeURIComponent(`uniquename eq '${uniqueNameOrId}'`)}`;
    }

    console.log(`[SolutionClient] GET /solutions?$select=${select}&$filter=uniquename eq '${uniqueNameOrId}'`);

    const response = await axios.get(url, { headers: await this.headers() });
    const solution = isGuid ? response.data : (response.data?.value || [])[0];

    if (!solution) {
      throw new Error(`Solution not found: ${uniqueNameOrId}`);
    }

    return {
      solutionId: solution.solutionid,
      uniqueName: solution.uniquename,
      friendlyName: solution.friendlyname,
      version: solution.version,
      description: solution.description || '',
      isManaged: solution.ismanaged,
      publisherId: solution._publisherid_value,
      installedOn: solution.installedon,
      modifiedOn: solution.modifiedon,
    };
  }

  // -----------------------------------------------------------------------
  // List Solution Components (returns array with componentTypeName)
  // -----------------------------------------------------------------------

  async listSolutionComponents(solutionId: string): Promise<any[]> {
    const url = `${this.apiBase}/solutioncomponents?$filter=${encodeURIComponent(`_solutionid_value eq '${solutionId}'`)}&$select=solutioncomponentid,componenttype,objectid,ismetadata&$orderby=componenttype asc`;
    console.log(`[SolutionClient] GET /solutioncomponents?$filter=_solutionid_value eq '${solutionId}'&$select=solutioncomponentid,componenttype,objectid,ismetadata&$orderby=componenttype asc`);

    const response = await axios.get(url, { headers: await this.headers() });
    const components = response.data?.value || [];
    console.log(`[SolutionClient] Solution ${solutionId}: ${components.length} components`);

    return components.map((c: any) => ({
      solutionComponentId: c.solutioncomponentid,
      componentType: c.componenttype,
      componentTypeName: componentTypeName(c.componenttype),
      objectId: c.objectid,
      isMetadata: c.ismetadata,
    }));
  }

  // -----------------------------------------------------------------------
  // Add Solution Component
  // -----------------------------------------------------------------------

  async addSolutionComponent(
    solutionUniqueName: string,
    componentId: string,
    componentType: number,
    addRequiredComponents: boolean = false
  ): Promise<any> {
    const url = `${this.apiBase}/AddSolutionComponent`;
    console.log(`[SolutionClient] POST /AddSolutionComponent (solution: ${solutionUniqueName}, type: ${componentType}, id: ${componentId})`);

    const body = {
      ComponentId: componentId,
      ComponentType: componentType,
      SolutionUniqueName: solutionUniqueName,
      AddRequiredComponents: addRequiredComponents,
    };

    try {
      const response = await axios.post(url, body, { headers: await this.headers() });
      console.log(`[SolutionClient] Component added to ${solutionUniqueName}`);

      return {
        success: true,
        solutionUniqueName,
        componentId,
        componentType,
        componentTypeName: componentTypeName(componentType),
        result: response.data,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[SolutionClient] addSolutionComponent failed (${status}): ${msg}`);
      throw new Error(`Failed to add component: ${status} — ${msg}`);
    }
  }

  // -----------------------------------------------------------------------
  // Export Solution (returns {fileName, sizeBytes, base64Content})
  // -----------------------------------------------------------------------

  async exportSolution(solutionName: string, managed: boolean = false): Promise<{ fileName: string; sizeBytes: number; base64Content: string }> {
    const url = `${this.apiBase}/ExportSolution`;
    console.log(`[SolutionClient] POST /ExportSolution (name: ${solutionName}, managed: ${managed})`);

    try {
      const body = {
        SolutionName: solutionName,
        Managed: managed,
      };

      const response = await axios.post(url, body, { headers: await this.headers() });
      const exportFile = response.data?.ExportSolutionFile || '';

      console.log(`[SolutionClient] Solution exported: ${solutionName} (managed: ${managed})`);

      return {
        fileName: `${solutionName}${managed ? '_managed' : ''}.zip`,
        sizeBytes: exportFile ? Math.round((exportFile.length * 3) / 4) : 0,
        base64Content: exportFile,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[SolutionClient] exportSolution failed (${status}): ${msg}`);
      throw new Error(`Failed to export solution: ${status} — ${msg}`);
    }
  }

  // -----------------------------------------------------------------------
  // Create Solution (v3.3.0 NEW)
  // -----------------------------------------------------------------------

  async createSolution(
    uniqueName: string,
    friendlyName: string,
    publisherId: string,
    version: string = '1.0.0.0',
    description: string = ''
  ): Promise<any> {
    const url = `${this.apiBase}/solutions`;
    console.log(`[SolutionClient] POST /solutions (name: ${uniqueName}, publisher: ${publisherId})`);

    try {
      const body: any = {
        uniquename: uniqueName,
        friendlyname: friendlyName,
        version,
        description,
        'publisherid@odata.bind': `/publishers(${publisherId})`,
      };

      const response = await axios.post(url, body, { headers: await this.headers() });

      // Dataverse returns 204 with OData-EntityId header, or 201 with body
      const solutionId =
        response.data?.solutionid ||
        response.headers['odata-entityid']?.match(/\(([^)]+)\)/)?.[1] ||
        'created';

      console.log(`[SolutionClient] Solution created: ${uniqueName} (${solutionId})`);

      return {
        success: true,
        solutionId,
        uniqueName,
        friendlyName,
        version,
        description,
        publisherId,
        message: `Solution '${friendlyName}' created successfully.`,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[SolutionClient] createSolution failed (${status}): ${msg}`);
      throw new Error(`Failed to create solution: ${status} — ${msg}`);
    }
  }

  // -----------------------------------------------------------------------
  // Delete Solution (v3.3.0 NEW)
  // -----------------------------------------------------------------------

  async deleteSolution(solutionId: string): Promise<any> {
    const url = `${this.apiBase}/solutions(${solutionId})`;
    console.log(`[SolutionClient] DELETE /solutions(${solutionId})`);

    try {
      await axios.delete(url, { headers: await this.headers() });
      console.log(`[SolutionClient] Solution deleted: ${solutionId}`);

      return {
        success: true,
        solutionId,
        message: `Solution ${solutionId} deleted successfully.`,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[SolutionClient] deleteSolution failed (${status}): ${msg}`);
      throw new Error(`Failed to delete solution: ${status} — ${msg}`);
    }
  }
}
