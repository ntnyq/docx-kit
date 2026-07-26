/**
 * MCP tool: get usage help for a specific plugin.
 *
 * @module mcp-server/tools/getPluginHelp
 */

import { findBuiltinPlugin } from '../plugins/catalog'
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
  const metadata = findBuiltinPlugin(plugin.name)

  return {
    name: plugin.name,
    description:
      metadata?.description
      ?? `Plugin: ${plugin.name}. Renders ${plugin.name} content in the document.`,
    usageExample:
      metadata?.usageExample
      ?? `{ type: "plugin", name: "${plugin.name}", options: { ... } }`,
  }
}
