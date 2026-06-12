/**
 * MCP server entry point (Node.js platform only).
 *
 * Requires @modelcontextprotocol/sdk and zod as optional peer deps.
 *
 * @module mcp
 */

export { docxSchemaResource } from './mcp-server/resources/schema'
export { createDocxToolDefinition } from './mcp-server/tools/createDocx'
export {
  applyTemplate,
  applyTemplateToolDefinition,
} from './mcp-server/tools/applyTemplate'
export {
  buildPluginHelp,
  getPluginHelpToolDefinition,
} from './mcp-server/tools/getPluginHelp'
export {
  buildPluginInfoList,
  listPluginsToolDefinition,
} from './mcp-server/tools/listPlugins'
export {
  validateSchema,
  validateSchemaToolDefinition,
} from './mcp-server/tools/validateSchema'
export {
  createDocxKitServer,
  RESOURCE_DEFINITIONS,
  TOOL_DEFINITIONS,
} from './mcp-server/index'
export {
  buildTemplateInfoList,
  listTemplatesToolDefinition,
} from './mcp-server/tools/listTemplates'
export type { PluginInfo } from './mcp-server/tools/listPlugins'
export type { PluginHelpInfo } from './mcp-server/tools/getPluginHelp'
export type {
  CreateDocumentInput,
  CreateDocumentOutput,
} from './mcp-server/tools/createDocx'
export type {
  ValidateSchemaOutput,
  ValidationError,
} from './mcp-server/tools/validateSchema'
export type {
  ApplyTemplateInput,
  ApplyTemplateOutput,
} from './mcp-server/tools/applyTemplate'
