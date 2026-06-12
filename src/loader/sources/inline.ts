/**
 * Inline plugin source loader.
 *
 * Wraps an already-loaded `DocxPlugin` instance as a
 * `PluginLoadResult`. This is an identity operation —
 * no loading or validation occurs.
 *
 * @module loader/sources/inline
 */

import type { DocxPlugin } from '../../types/plugin'

/**
 * Load an inline (already-in-memory) plugin.
 *
 * @param plugin - — The plugin instance
 * @returns A load result with the same plugin and `null` manifest
 */
export function loadInlinePlugin(plugin: DocxPlugin): {
  manifest: null
  plugin: DocxPlugin
} {
  return { manifest: null, plugin }
}
