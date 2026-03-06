export interface TokenInfo {
  accessToken: string;
  expiresAt: number;
  scope: string;
}

export interface ToolResult {
  [key: string]: unknown;
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, any>;
    required?: string[];
  };
  handler: (params: Record<string, any>) => Promise<ToolResult>;
}
