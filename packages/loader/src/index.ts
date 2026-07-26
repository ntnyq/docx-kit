/**
 * @docxkit/loader — Plugin loader with platform-specific sources.
 *
 * Re-exports the core `PluginLoader` and `createPluginLoader` from
 * `@docxkit/core`, plus source-specific standalone loader functions.
 *
 * For platform-specific source loaders (npm, url), use the subpath exports:
 * - `@docxkit/loader/browser` — includes `loadUrlPlugin`
 * - `@docxkit/loader/node` — includes `loadNpmPlugin`
 *
 * @packageDocumentation
 */

export { loadLocalPlugin } from './sources/local'

// Cross-platform source loaders
export { loadInlinePlugin } from './sources/inline'

// Re-export core loader primitives
export {
  createPluginLoader,
  PluginLoader,
  validateManifest,
} from '@docxkit/core'
export type {
  PluginLoaderOptions,
  PluginLoadResult,
  PluginManifest,
  PluginManifestAuthorizer,
  PluginSecurityPolicy,
  PluginSource,
} from '@docxkit/core'
