/**
 * MCP tool: list available docx-kit plugins.
 *
 * @module mcp-server/tools/listPlugins
 */

import type { DocxPlugin } from '../../types/plugin'

/**
 * Plugin info returned by the list_plugins tool.
 */
export interface PluginInfo {
  /** Plugin description (derived from name and render signature). */
  description: string
  /** Plugin name. */
  name: string
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
 * @param plugins - — Registered plugins
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
