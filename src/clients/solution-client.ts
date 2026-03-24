/**
 * Solution Client — Power Automate MCP (Dataverse Web API)
 * v3.2.0: Initial (list, get, components, add-component, export)
 * v3.3.0: Added createSolution, deleteSolution
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

  constructor(dataverseUrl: string | undefined, tokenProvider: TokenProvider) {
    if (!dataverseUrl) {
      this._configured = false;
      this.dataverseUrl = '';
      this.scope = '';
      this.apiBase = '';
      this.tokenProvider = tokenProvider;
      console.log('[SolutionClient] Dataverse not configured (DATAVERSE_URL missing)');
      return;
    }

    this.dataverseUrl = dataverseUrl.startsWith('https://')
      ? dataverseUrl
      : `https://${dataverseUrl}`;
    this.scope = `${this.dataverseUrl}/.default`;
    this.apiBase = `${this.dataverseUrl}/api/data/v9.2`;
    this.tokenProvider = tokenProvider;
    this._configured = true;

    console.log(`[SolutionClient] Dataverse configured: ${this.dataverseUrl}`);
    console.log(`[SolutionClient]   Scope: ${this.scope}`);
  }

  get configured(): boolean {
    return this._configured;
  }

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

  /**
   * List Dataverse solutions.
   */
  async listSolutions(unmanagedOnly: boolean = true): Promise<any> {
    const select = 'solutionid,uniquename,friendlyname,version,ismanaged,installedon,modifiedon,description,_publisherid_value';
    let filter = 'isvisible eq true';
    if (unmanagedOnly) filter += ' and ismanaged eq false';

    const url = `${this.apiBase}/solutions?$select=${select}&$filter=${encodeURIComponent(filter)}&$orderby=friendlyname asc`;
    console.log(`[SolutionClient] GET /solutions?$select=${select}&$filter=${filter}&$orderby=friendlyname asc`);

    const response = await axios.get(url, { headers: await this.headers() });
    const solutions = response.data?.value || [];
    console.log(`[SolutionClient] Found ${solutions.length} solutions`);

    return {
      count: solutions.length,
      filter: unmanagedOnly ? 'unmanaged only' : 'all visible',
      solutions: solutions.map((s: any) => ({
        solutionId: s.solutionid,
        uniqueName: s.uniquename,
        friendlyName: s.friendlyname,
        version: s.version,
        isManaged: s.ismanaged,
        description: s.description || '',
        publisherId: s._publisherid_value,
        installedOn: s.installedon,
        modifiedOn: s.modifiedon,
      })),
    };
  }

  /**
   * Get a solution by unique name or GUID.
   */
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

  /**
   * List components in a solution.
   */
  async listSolutionComponents(solutionId: string): Promise<any> {
    const url = `${this.apiBase}/solutioncomponents?$filter=${encodeURIComponent(`_solutionid_value eq '${solutionId}'`)}&$select=solutioncomponentid,componenttype,objectid,ismetadata&$orderby=componenttype asc`;
    console.log(`[SolutionClient] GET /solutioncomponents?$filter=_solutionid_value eq '${solutionId}'&$select=solutioncomponentid,componenttype,objectid,ismetadata&$orderby=componenttype asc`);

    const response = await axios.get(url, { headers: await this.headers() });
    const components = response.data?.value || [];
    console.log(`[SolutionClient] Solution ${solutionId}: ${components.length} components`);

    // Group by component type
    const byType: Record<string, number> = {};
    const mapped = components.map((c: any) => {
      const typeName = componentTypeName(c.componenttype);
      byType[typeName] = (byType[typeName] || 0) + 1;
      return {
        solutionComponentId: c.solutioncomponentid,
        componentType: c.componenttype,
        componentTypeName: typeName,
        objectId: c.objectid,
        isMetadata: c.ismetadata,
      };
    });

    return {
      solutionId,
      count: mapped.length,
      componentsByType: byType,
      components: mapped,
    };
  }

  /**
   * Add a component to a solution.
   */
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

  /**
   * Export a solution as a ZIP (base64).
   */
  async exportSolution(solutionName: string, managed: boolean = false): Promise<any> {
    const url = `${this.apiBase}/ExportSolution`;
    console.log(`[SolutionClient] POST /ExportSolution (name: ${solutionName}, managed: ${managed})`);

    try {
      const body = {
        SolutionName: solutionName,
        Managed: managed,
      };

      const response = await axios.post(url, body, { headers: await this.headers() });
      const exportFile = response.data?.ExportSolutionFile;

      console.log(`[SolutionClient] Solution exported: ${solutionName} (managed: ${managed})`);

      return {
        success: true,
        solutionName,
        managed,
        fileName: `${solutionName}${managed ? '_managed' : ''}.zip`,
        fileSize: exportFile ? Math.round((exportFile.length * 3) / 4) : 0,
        base64Content: exportFile,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[SolutionClient] exportSolution failed (${status}): ${msg}`);
      throw new Error(`Failed to export solution: ${status} — ${msg}`);
    }
  }

  /**
   * Create a new Dataverse solution.
   * v3.3.0: NEW tool
   */
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

      const response = await axios.post(url, body, {
        headers: await this.headers(),
      });

      // Dataverse returns 204 with OData-EntityId header, or 201 with body
      const solutionId =
        response.data?.solutionid ||
        response.headers['odata-entityid']?.match(/\(([^)]+)\)/)?.[1] ||
        'unknown';

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

  /**
   * Delete a Dataverse solution by ID.
   * v3.3.0: NEW tool
   */
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
