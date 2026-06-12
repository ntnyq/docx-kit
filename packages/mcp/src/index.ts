/**
 * docx-kit MCP server — AI agent integration via
 * Model Context Protocol.
 *
 * Provides tools for document creation, validation,
 * plugin discovery, template application, and schema resources.
 *
 * The server uses `@modelcontextprotocol/sdk` (optional peer dep)
 * for MCP protocol handling. If the SDK is not installed,
 * individual tool/resource definitions and execution logic
 * can still be used programmatically.
 *
 * @module mcp-server
 */

import { docxSchemaResource } from './resources/schema'
import {
  applyTemplate,
  applyTemplateToolDefinition,
} from './tools/applyTemplate'
import { createDocxToolDefinition } from './tools/createDocx'
import { getPluginHelpToolDefinition } from './tools/getPluginHelp'
import { listPluginsToolDefinition } from './tools/listPlugins'
import {
  buildTemplateInfoList,
  listTemplatesToolDefinition,
} from './tools/listTemplates'
import {
  validateSchema,
  validateSchemaToolDefinition,
} from './tools/validateSchema'

/**
 * All MCP tool definitions for docx-kit.
 */
export const TOOL_DEFINITIONS = [
  applyTemplateToolDefinition,
  createDocxToolDefinition,
  getPluginHelpToolDefinition,
  listPluginsToolDefinition,
  listTemplatesToolDefinition,
  validateSchemaToolDefinition,
] as const

/**
 * All MCP resource definitions for docx-kit.
 */
export const RESOURCE_DEFINITIONS = [docxSchemaResource] as const

/**
 * Create a docx-kit MCP server.
 *
 * Uses `@modelcontextprotocol/sdk` to create a fully compliant
 * MCP server with all docx-kit tools and resources registered.
 *
 * Requires `@modelcontextprotocol/sdk` to be installed.
 * Connect via `StdioServerTransport` for CLI usage or
 * `StreamableHTTPServerTransport` for HTTP access.
 *
 * @returns An MCP server instance ready to connect
 *
 * @example
 * ```ts
 * import { createDocxKitServer } from 'docx-kit/mcp'
 * import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
 *
 * const server = createDocxKitServer()
 * const transport = new StdioServerTransport()
 * await server.connect(transport)
 * ```
 */
export async function createDocxKitServer(): Promise<unknown> {
  // Dynamic import — SDK is an optional peer dependency
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js')
  const { z } = await import('zod')

  const server = new McpServer({
    name: 'docx-kit',
    version: '0.2.0',
  })

  // Tool: create_document
  server.registerTool(
    'create_document',
    {
      description: 'Create a new .docx document from a docx-kit JSON schema',
      inputSchema: {
        outputPath: z.string().describe('File path for the output .docx file'),
        schema: z
          .record(z.string(), z.unknown())
          .describe('A docx-kit DocxSchema JSON object'),
      },
    },
    async (input: { outputPath: string; schema: Record<string, unknown> }) => {
      // For MCP, we return a text description (actual file saving
      // must be handled by the caller with renderDocx() + save())
      const docSchema = input.schema
      const validation = validateSchema(docSchema as never)

      if (!validation.valid) {
        return {
          isError: true,
          content: [
            {
              text: `Schema validation failed:\n${validation.errors.map(e => `  ${e.path}: ${e.message}`).join('\n')}`,
              type: 'text' as const,
            },
          ],
        }
      }

      return {
        content: [
          {
            text: `Document schema is valid. Use renderDocx() to generate the document and save to: ${input.outputPath}`,
            type: 'text' as const,
          },
        ],
      }
    },
  )

  // Tool: validate_schema
  server.registerTool(
    'validate_schema',
    {
      description: 'Validate a docx-kit JSON schema for correctness',
      inputSchema: {
        schema: z
          .record(z.string(), z.unknown())
          .describe('The DocxSchema JSON object to validate'),
      },
    },
    async (input: { schema: Record<string, unknown> }) => {
      const result = validateSchema(input.schema as never)
      return {
        content: [
          {
            type: 'text' as const,
            text: result.valid
              ? 'Schema is valid ✓'
              : `Validation errors:\n${result.errors.map(e => `  ${e.path}: ${e.message}`).join('\n')}`,
          },
        ],
      }
    },
  )

  // Tool: list_plugins
  server.registerTool(
    'list_plugins',
    {
      description: 'List all available docx-kit plugins',
      inputSchema: {
        filter: z.string().optional().describe('Optional name filter pattern'),
      },
    },
    async (input: { filter?: string }) => {
      // Built-in plugin names (no runtime plugin list available in MCP context)
      const builtInNames = [
        'callout',
        'codeBlock',
        'coverPage',
        'dataTable',
        'echarts',
        'meetingMinutes',
        'pageNumber',
        'propertyTable',
        'qrcode',
        'signatureBlock',
        'timeline',
        'watermark',
      ]
      const infoList = builtInNames.map(name => ({
        description: `Built-in docx-kit plugin: ${name}`,
        name,
      }))
      const filtered = input.filter
        ? infoList.filter(p =>
            p.name.toLowerCase().includes(input.filter!.toLowerCase()),
          )
        : infoList

      return {
        content: [
          {
            text: JSON.stringify(filtered, null, 2),
            type: 'text' as const,
          },
        ],
      }
    },
  )

  // Tool: get_plugin_help
  server.registerTool(
    'get_plugin_help',
    {
      description: 'Get usage help for a specific docx-kit plugin',
      inputSchema: {
        pluginName: z.string().describe('The plugin name'),
      },
    },
    async (input: { pluginName: string }) => {
      // Create a minimal DocxPlugin reference for help
      const helpInfo: Record<string, unknown> = {
        description: `Built-in docx-kit plugin: ${input.pluginName}`,
        name: input.pluginName,
        usageExample: `{ type: "plugin", name: "${input.pluginName}", options: { ... } }`,
      }

      const usageExamples: Record<string, string> = {
        pageNumber: '{ type: "plugin", name: "pageNumber", options: {} }',
        callout:
          '{ type: "plugin", name: "callout", options: { title: "Note", variant: "info", text: "Important message" } }',
        codeBlock:
          '{ type: "plugin", name: "codeBlock", options: { code: "console.log(\'hello\')", language: "javascript" } }',
        coverPage:
          '{ type: "plugin", name: "coverPage", options: { title: "Annual Report", author: "John Doe", date: "2026-06-12" } }',
        dataTable:
          '{ type: "plugin", name: "dataTable", options: { columns: [...], data: [...], striped: true } }',
        echarts:
          '{ type: "plugin", name: "echarts", options: { option: { ...echarts config... }, width: 400, height: 300 } }',
        meetingMinutes:
          '{ type: "plugin", name: "meetingMinutes", options: { title: "Team Meeting", date: "2026-06-12", attendees: ["Alice", "Bob"] } }',
        propertyTable:
          '{ type: "plugin", name: "propertyTable", options: { items: [{ key: "Name", value: "Alice" }] } }',
        qrcode:
          '{ type: "plugin", name: "qrcode", options: { text: "https://example.com", width: 100 } }',
        signatureBlock:
          '{ type: "plugin", name: "signatureBlock", options: { parties: [{ name: "Alice", role: "Manager" }] } }',
        timeline:
          '{ type: "plugin", name: "timeline", options: { events: [{ date: "2026-01", title: "Kickoff" }] } }',
        watermark:
          '{ type: "plugin", name: "watermark", options: { text: "CONFIDENTIAL", opacity: 0.3 } }',
      }

      if (usageExamples[input.pluginName]) {
        helpInfo.usageExample = usageExamples[input.pluginName]
      }

      return {
        content: [
          {
            text: JSON.stringify(helpInfo, null, 2),
            type: 'text' as const,
          },
        ],
      }
    },
  )

  // Tool: list_templates
  server.registerTool(
    'list_templates',
    {
      description: 'List all available docx-kit AI templates',
    },
    async () => {
      const templates = buildTemplateInfoList()
      return {
        content: [
          {
            text: JSON.stringify(templates, null, 2),
            type: 'text' as const,
          },
        ],
      }
    },
  )

  // Tool: apply_template
  server.registerTool(
    'apply_template',
    {
      description: 'Generate a document schema from an AI template + data',
      inputSchema: {
        data: z
          .record(z.string(), z.unknown())
          .describe('Template parameter values'),
        template: z
          .string()
          .describe('Template name (report, invoice, resume, letter)'),
      },
    },
    async (input: { data: Record<string, unknown>; template: string }) => {
      const result = applyTemplate(input.template, input.data)
      if (!result) {
        return {
          isError: true,
          content: [
            {
              text: `Template "${input.template}" not found. Available: report, invoice, resume, letter`,
              type: 'text' as const,
            },
          ],
        }
      }
      return {
        content: [
          {
            text: JSON.stringify(result.schema, null, 2),
            type: 'text' as const,
          },
        ],
      }
    },
  )

  // Resource: docx-kit://schema
  server.registerResource(
    'docx-kit-schema',
    'docx-kit://schema',
    {},
    async (uri: URL) => {
      return {
        contents: [
          {
            mimeType: 'application/json',
            text: JSON.stringify(docxSchemaResource.schema, null, 2),
            uri: uri.href,
          },
        ],
      }
    },
  )

  return server
}

export { docxSchemaResource } from './resources/schema'
export { createDocxToolDefinition } from './tools/createDocx'
// Re-export individual tool/resource modules for programmatic use
export {
  applyTemplate,
  applyTemplateToolDefinition,
} from './tools/applyTemplate'
export {
  buildPluginHelp,
  getPluginHelpToolDefinition,
} from './tools/getPluginHelp'
export {
  buildPluginInfoList,
  listPluginsToolDefinition,
} from './tools/listPlugins'
export {
  validateSchema,
  validateSchemaToolDefinition,
} from './tools/validateSchema'
export {
  buildTemplateInfoList,
  listTemplatesToolDefinition,
} from './tools/listTemplates'
export type { PluginInfo } from './tools/listPlugins'
export type { PluginHelpInfo } from './tools/getPluginHelp'
export type {
  CreateDocumentInput,
  CreateDocumentOutput,
} from './tools/createDocx'
export type {
  ValidateSchemaOutput,
  ValidationError,
} from './tools/validateSchema'
export type {
  ApplyTemplateInput,
  ApplyTemplateOutput,
} from './tools/applyTemplate'
