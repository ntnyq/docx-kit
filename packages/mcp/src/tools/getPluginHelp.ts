/**
 * MCP tool: get usage help for a specific plugin.
 *
 * @module mcp-server/tools/getPluginHelp
 */

import type { DocxPlugin } from '@docxkit/core'

/**
 * Plugin help info returned by get_plugin_help.
 */
export interface PluginHelpInfo {
  /** Detailed description of the plugin. */
  description: string
  /** Plugin name. */
  name: string
  /** Example usage in JSON DSL. */
  usageExample: string
}

/**
 * MCP tool definition for `get_plugin_help`.
 *
 * Returns usage help and option schema for a named plugin.
 */
export const getPluginHelpToolDefinition = {
  name: 'get_plugin_help',
  description:
    'Get usage help and option details for a specific docx-kit plugin.',
  inputSchema: {
    required: ['pluginName'],
    type: 'object',
    properties: {
      pluginName: {
        description: 'The name of the plugin to get help for',
        type: 'string',
      },
    },
  },
}

/**
 * Build help info for a plugin.
 *
 * Provides a usage example based on the plugin name,
 * referencing a built-in map of known plugins.
 *
 * @param plugin - — The DocxPlugin to build help for
 * @returns PluginHelpInfo with description, name, and usage example
 */
export function buildPluginHelp(plugin: DocxPlugin): PluginHelpInfo {
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

  return {
    description: `Built-in docx-kit plugin: ${plugin.name}. Renders ${plugin.name} content in the document.`,
    name: plugin.name,
    usageExample:
      usageExamples[plugin.name]
      ?? `{ type: "plugin", name: "${plugin.name}", options: { ... } }`,
  }
}
