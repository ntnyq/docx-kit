import type { ErrorCode } from './catalog'

/**
 * Runtime ownership and documentation source, without translated prose.
 */
export interface ErrorMetadata {
  packageName: string
  source: string
  category:
    | 'export'
    | 'image'
    | 'manifest'
    | 'node'
    | 'plugin'
    | 'preview'
    | 'style'
    | 'table'
}

/**
 * Every catalog entry has an owner and a portable repository source link.
 */
export const ERROR_METADATA = {
  EXPORT_FAILED: {
    category: 'export',
    packageName: '@docxkit/core',
    source: 'packages/core/src/renderer/pack.ts',
  },
  HEADER_FOOTER_INVALID_CONTENT: {
    category: 'node',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/compileDocument.ts',
  },
  IMAGE_INVALID_DATA: {
    category: 'image',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/nodes/compileImage.ts',
  },
  MANIFEST_INVALID: {
    category: 'manifest',
    packageName: '@docxkit/core',
    source: 'packages/core/src/loader/manifest.ts',
  },
  MANIFEST_MISSING: {
    category: 'manifest',
    packageName: '@docxkit/loader',
    source: 'packages/loader/src/sources/npm.ts',
  },
  MICROSOFT_URL_REQUIRED: {
    category: 'preview',
    packageName: '@docxkit/renderer',
    source: 'packages/renderer/src/render.ts',
  },
  PLUGIN_DEPENDENCY_MISSING: {
    category: 'plugin',
    packageName: '@docxkit/errors',
    source: 'packages/errors/src/catalog.ts',
  },
  PLUGIN_LOAD_FAILED: {
    category: 'plugin',
    packageName: '@docxkit/loader',
    source: 'packages/loader/src/utils.ts',
  },
  PLUGIN_NOT_REGISTERED: {
    category: 'plugin',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/nodes/compilePlugin.ts',
  },
  PLUGIN_RENDER_FAILED: {
    category: 'plugin',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/nodes/compilePlugin.ts',
  },
  PLUGIN_VERSION_MISMATCH: {
    category: 'plugin',
    packageName: '@docxkit/core',
    source: 'packages/core/src/loader/PluginLoader.ts',
  },
  PREVIEW_FETCH_FAILED: {
    category: 'preview',
    packageName: '@docxkit/renderer',
    source: 'packages/renderer/src/normalize.ts',
  },
  PREVIEW_INPUT_INVALID: {
    category: 'preview',
    packageName: '@docxkit/renderer',
    source: 'packages/renderer/src/normalize.ts',
  },
  PREVIEW_RENDER_FAILED: {
    category: 'preview',
    packageName: '@docxkit/renderer',
    source: 'packages/renderer/src/render.ts',
  },
  STYLE_CIRCULAR_EXTENDS: {
    category: 'style',
    packageName: '@docxkit/core',
    source: 'packages/core/src/style/normalizeStyle.ts',
  },
  STYLE_UNKNOWN_CLASS: {
    category: 'style',
    packageName: '@docxkit/core',
    source: 'packages/core/src/style/normalizeStyle.ts',
  },
  TABLE_INVALID_COLUMNS: {
    category: 'table',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/nodes/compileTable.ts',
  },
  UNKNOWN_NODE_TYPE: {
    category: 'node',
    packageName: '@docxkit/core',
    source: 'packages/core/src/compiler/registry.ts',
  },
} as const satisfies Record<ErrorCode, ErrorMetadata>
