/**
 * docx-kit Plugin Loader — shared exports.
 *
 * Re-exported by both `docx-kit/loader` (browser) and
 * `docx-kit/loader/node`. For platform-specific source loaders,
 * use the platform entry points.
 *
 * @module docx-kit/loader
 * @packageDocumentation
 */

// Manifest validation
export { validateManifest } from './loader/manifest'

export { loadLocalPlugin } from './loader/sources/local'

// Source loaders (cross-platform)
export { loadInlinePlugin } from './loader/sources/inline'
// Core loader
export { createPluginLoader, PluginLoader } from './loader/PluginLoader'

export type { PluginManifest } from './loader/manifest'

// Types
export type {
  PluginLoadResult,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'
