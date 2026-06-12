/**
 * Plugin registry entry point — re-exports all registry APIs.
 *
 * @module registry
 */

export { getPlugin, searchPlugins } from './registry/PluginSearch'
export { createPluginRegistry, PluginRegistry } from './registry/PluginRegistry'
export type {
  NpmSearchResult,
  QualityScore,
  RegistryPluginEntry,
} from './registry/types'
