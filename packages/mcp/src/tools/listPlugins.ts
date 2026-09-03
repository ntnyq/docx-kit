/**
 * MCP tool: list available docx-kit plugins.
 *
 * @module mcp-server/tools/listPlugins
 */

import { BUILTIN_PLUGIN_CATALOG } from '../plugins/catalog'
import type { DocxPlugin } from '@docxkit/core'

/**
 * Plugin info returned by the list_plugins tool.
 */
export interface PluginInfo {
  /**
   * Plugin description (derived from name and render signature).
   */
  description: string
  /**
   * Plugin name.
   */
  name: string
}

/**
 * Build plugin info for every built-in plugin in the canonical catalog.
 *
 * @param filter - Optional case-insensitive substring used to match plugin names
 * @returns Matching built-in plugin names and descriptions in catalog order
 */
export function buildBuiltinPluginInfoList(filter?: string): PluginInfo[] {
  const normalizedFilter = filter?.toLowerCase()
  return BUILTIN_PLUGIN_CATALOG.filter(
    plugin =>
      !normalizedFilter || plugin.name.toLowerCase().includes(normalizedFilter),
  ).map(plugin => ({
    description: plugin.description,
    name: plugin.name,
  }))
}

/**
 * MCP tool definition for `list_plugins`.
 *
 * Returns metadata for all registered plugins.
 */
export const listPluginsToolDefinition = {
  name: 'list_plugins',
  description:
    'List all available docx-kit plugins with their names and descriptions.',
  inputSchema: {
    type: 'object',
    properties: {
      filter: {
        description: 'Optional name filter pattern',
        type: 'string',
      },
    },
  },
}

/**
 * Build plugin info from a list of DocxPlugin instances.
 *
 * Optionally filters by a case-insensitive name substring.
 *
 * @param plugins - — Registered plugins
 * @param filter - — Optional name filter pattern (case-insensitive)
 * @returns Array of PluginInfo objects
 */
export function buildPluginInfoList(
  plugins: DocxPlugin[],
  filter?: string,
): PluginInfo[] {
  const allInfo: PluginInfo[] = plugins.map(p => ({
    description: `Built-in docx-kit plugin: ${p.name}`,
    name: p.name,
  }))

  if (filter) {
    return allInfo.filter(p =>
      p.name.toLowerCase().includes(filter.toLowerCase()),
    )
  }

  return allInfo
}
