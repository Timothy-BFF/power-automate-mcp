/**
 * Connection Client — Power Automate MCP
 * v3.3.0: Fixed listConnections endpoint (admin → delegated), added createConnection
 *
 * Bug fix: Jose Sanchez reported pa-list-connections returning 403.
 * Root cause: Was using /scopes/admin/environments/{envId}/connections
 * Fix: Now uses /environments/{envId}/connections (delegated path)
 */
import axios from 'axios';

const FLOW_API = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple';

interface TokenProvider {
  getToken(scope: string): Promise<string>;
}

interface UserTokenProvider {
  getAccessToken(userId: string): Promise<string | null>;
}

export class ConnectionClient {
  private tokenProvider: TokenProvider;
  private userTokenProvider?: UserTokenProvider;

  constructor(tokenProvider: TokenProvider, userTokenProvider?: UserTokenProvider) {
    this.tokenProvider = tokenProvider;
    this.userTokenProvider = userTokenProvider;
  }

  private async getFlowToken(): Promise<string> {
    return this.tokenProvider.getToken('https://service.flow.microsoft.com/.default');
  }

  private async getUserToken(userId?: string): Promise<string> {
    if (userId && this.userTokenProvider) {
      const token = await this.userTokenProvider.getAccessToken(userId);
      if (token) return token;
    }
    return this.getFlowToken();
  }

  /**
   * List connections in an environment.
   * v3.3.0 FIX: Changed from admin-scoped to delegated path.
   * OLD (broken): /scopes/admin/environments/{envId}/connections
   * NEW (fixed):  /environments/{envId}/connections
   */
  async listConnections(environmentId: string, userId?: string): Promise<any> {
    const token = await this.getUserToken(userId);
    const url = `${FLOW_API}/environments/${environmentId}/connections`;
    console.log(`[ConnectionClient] GET ${url}`);

    try {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      const connections = response.data?.value || [];
      console.log(`[ConnectionClient] Found ${connections.length} connections`);

      return {
        count: connections.length,
        connections: connections.map((conn: any) => ({
          name: conn.name,
          id: conn.name,
          displayName: conn.properties?.displayName || conn.name,
          connectorName: conn.properties?.apiId?.split('/').pop() || 'unknown',
          status: conn.properties?.statuses?.[0]?.status || 'unknown',
          createdTime: conn.properties?.createdTime,
          lastModifiedTime: conn.properties?.lastModifiedTime,
          createdBy: conn.properties?.createdBy?.displayName,
          environment: environmentId,
        })),
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[ConnectionClient] listConnections failed (${status}): ${msg}`);
      throw new Error(`Failed to list connections: ${status} — ${msg}`);
    }
  }

  /**
   * Get details of a specific connection.
   */
  async getConnection(environmentId: string, connectionId: string, userId?: string): Promise<any> {
    const token = await this.getUserToken(userId);
    const url = `${FLOW_API}/environments/${environmentId}/connections/${connectionId}`;
    console.log(`[ConnectionClient] GET ${url}`);

    try {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      const conn = response.data;
      return {
        name: conn.name,
        id: conn.name,
        displayName: conn.properties?.displayName || conn.name,
        connectorName: conn.properties?.apiId?.split('/').pop() || 'unknown',
        connectorId: conn.properties?.apiId,
        status: conn.properties?.statuses?.[0]?.status || 'unknown',
        createdTime: conn.properties?.createdTime,
        lastModifiedTime: conn.properties?.lastModifiedTime,
        createdBy: conn.properties?.createdBy,
        environment: environmentId,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[ConnectionClient] getConnection failed (${status}): ${msg}`);
      throw new Error(`Failed to get connection ${connectionId}: ${status} — ${msg}`);
    }
  }

  /**
   * List available connectors in an environment.
   */
  async listConnectors(environmentId: string): Promise<any> {
    const token = await this.getFlowToken();
    const url = `${FLOW_API}/environments/${environmentId}/apis?$top=250`;
    console.log(`[ConnectionClient] GET connectors for ${environmentId}`);

    try {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      const connectors = response.data?.value || [];
      console.log(`[ConnectionClient] Found ${connectors.length} connectors`);

      return {
        count: connectors.length,
        connectors: connectors.map((c: any) => ({
          name: c.name,
          displayName: c.properties?.displayName || c.name,
          description: c.properties?.description || '',
          tier: c.properties?.tier || 'unknown',
          iconUri: c.properties?.iconUri,
        })),
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[ConnectionClient] listConnectors failed (${status}): ${msg}`);
      throw new Error(`Failed to list connectors: ${status} — ${msg}`);
    }
  }

  /**
   * Create a new connection (requires user delegated token).
   * v3.3.0: NEW tool
   */
  async createConnection(
    environmentId: string,
    connectorId: string,
    userId: string,
    connectionParameters?: Record<string, any>
  ): Promise<any> {
    if (!this.userTokenProvider) {
      throw new Error('User authentication required. Use pa-auth-start to begin device login.');
    }

    const token = await this.userTokenProvider.getAccessToken(userId);
    if (!token) {
      throw new Error(`Not authenticated. Use pa-auth-start to begin device login for ${userId}.`);
    }

    // Normalize connector API ID
    const apiId = connectorId.startsWith('/providers/')
      ? connectorId
      : `/providers/Microsoft.ProcessSimple/environments/${environmentId}/apis/${connectorId}`;

    const url = `${FLOW_API}/environments/${environmentId}/connections`;
    console.log(`[ConnectionClient] POST ${url} (connector: ${connectorId}, user: ${userId})`);

    try {
      const body = {
        properties: {
          apiId,
          connectionParameters: connectionParameters || {},
        },
      };

      const response = await axios.post(url, body, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      const conn = response.data;
      console.log(`[ConnectionClient] Connection created: ${conn.name}`);

      return {
        name: conn.name,
        id: conn.name,
        displayName: conn.properties?.displayName || conn.name,
        connectorName: connectorId,
        status: conn.properties?.statuses?.[0]?.status || 'created',
        createdTime: conn.properties?.createdTime,
        consentLink: conn.properties?.connectionParameters?.token?.oAuthSettings?.redirectUrl,
        message: `Connection created successfully. ${
          conn.properties?.statuses?.[0]?.status === 'Connected'
            ? 'Connection is active.'
            : 'You may need to authorize this connection in the Power Automate portal.'
        }`,
      };
    } catch (error: any) {
      const status = error.response?.status;
      const msg = error.response?.data?.error?.message || error.message;
      console.error(`[ConnectionClient] createConnection failed (${status}): ${msg}`);
      throw new Error(`Failed to create connection: ${status} — ${msg}`);
    }
  }
}
