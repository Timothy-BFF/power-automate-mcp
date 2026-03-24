/**
 * MCP SkillEngine — Power Automate MCP v1.0.0
 *
 * Central registration point for all MCP skills (prompts + resources).
 * Called from index.ts during server initialization.
 *
 * Skills provide operational knowledge that agents read BEFORE calling tools,
 * reducing trial-and-error and preventing common errors such as:
 *
 * - Parameter naming mismatches (snake_case vs camelCase)
 *   → Prevented by: parameter-conventions resource
 *
 * - OAuth connection "Error" confusion (expected status, not a bug)
 *   → Prevented by: connection-lifecycle resource + workflow-create-connection prompt
 *
 * - Skipping mandatory flow creation steps (causes orphaned flows)
 *   → Prevented by: workflow-create-flow prompt
 *
 * - Missing authentication before write operations
 *   → Prevented by: workflow-auth prompt
 *
 * Architecture:
 *   prompts.ts  → 3 workflow prompts (step-by-step guides)
 *   resources.ts → 3 knowledge resources (reference documents)
 *   register.ts → this file (aggregator)
 *
 * Usage in index.ts:
 *   import { registerSkills } from './skills/register.js';
 *   registerSkills(mcpServer);
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerPrompts } from './prompts.js';
import { registerResources } from './resources.js';

export function registerSkills(server: McpServer): void {
  registerPrompts(server);
  registerResources(server);
  console.log('[Skills] SkillEngine v1.0.0 initialized — 3 prompts + 3 resources');
}
