/**
 * Plugin registry entry point — re-exports all registry APIs.
 *
 * @module registry
 */

export { getPlugin, searchPlugins } from './PluginSearch'
export { createPluginRegistry, PluginRegistry } from './PluginRegistry'
export type {
  NpmSearchResult,
  QualityScore,
  RegistryPluginEntry,
} from './types'
