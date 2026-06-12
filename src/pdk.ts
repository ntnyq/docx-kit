/**
 * docx-kit PDK (Plugin Development Kit)
 *
 * Tools for building, testing, and validating docx-kit plugins.
 *
 * @module docx-kit/pdk
 * @packageDocumentation
 */

export {
  assertPluginDefined,
  assertRendersChildType,
  assertRendersParagraph,
  createPluginTestContext,
  renderPlugin,
} from './pdk/test-harness'

// Re-export loader types for convenience
export type {
  PluginLoadResult,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'
