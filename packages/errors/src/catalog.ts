import { defineDiagnostics } from 'nostics'

/**
 * Selects a repair for plugin loading without weakening security policy.
 */
export type PluginLoadReason =
  'environment' | 'export' | 'factory' | 'import' | 'security'

const pluginLoadFixes: Record<PluginLoadReason, string> = {
  environment:
    'Use the platform-specific loader, or register an inline plugin instance.',
  export:
    'Export a DocxPlugin instance or a factory returning one from the plugin module.',
  factory:
    'Inspect the original cause and correct the plugin factory or its dependencies.',
  import:
    'Check the plugin entry, installed dependencies, and network access for this platform.',
  security:
    'Review the plugin source and permissions against your allowLoad/allowExecute policy; keep untrusted code blocked.',
}

/**
 * Typed, silent Nostics definitions. Public legacy codes are derived below.
 */
export const diagnosticDefinitions = {
  DOCXKIT_EXPORT_FAILED: {
    fix: 'Check the original cause, document resources, and the selected export platform.',
    why(params: { format: 'base64' | 'blob' | 'buffer' }) {
      return `Failed to pack document to ${params.format}`
    },
  },
  DOCXKIT_HEADER_FOOTER_INVALID_CONTENT: {
    fix: 'Return Paragraph or Table instances from plugins used in headers and footers.',
    why: 'Header and footer plugins must render paragraphs or tables',
  },
  DOCXKIT_IMAGE_INVALID_DATA: {
    fix: 'Supply non-empty supported image bytes or a data URL; use docx-kit/node or resolveImage for paths.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_MANIFEST_INVALID: {
    fix: 'Correct the named manifest field, semver range, or safe relative entry path in docx-kit.plugin.json.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_MANIFEST_MISSING: {
    fix: 'Include docx-kit.plugin.json at the plugin package root and verify the resolved path or URL.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_MICROSOFT_URL_REQUIRED: {
    fix: 'Pass a publicly accessible DOCX URL to the Microsoft renderer, or use the local docx-preview renderer.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_PLUGIN_DEPENDENCY_MISSING: {
    fix: 'Install the dependency and a version compatible with the plugin manifest.',
    why(params: { dependency: string }) {
      return `Plugin dependency missing: ${params.dependency}`
    },
  },
  DOCXKIT_PLUGIN_LOAD_FAILED: {
    fix(params: { detail: string; reason: PluginLoadReason }) {
      return pluginLoadFixes[params.reason]
    },
    why(params: { detail: string; reason: PluginLoadReason }) {
      return params.detail
    },
  },
  DOCXKIT_PLUGIN_NOT_REGISTERED: {
    fix: 'Register the plugin instance with .use(pluginFactory()) before exporting, and match its runtime name.',
    why(params: { pluginName: string }) {
      return `Plugin not registered: ${params.pluginName}`
    },
  },
  DOCXKIT_PLUGIN_RENDER_FAILED: {
    fix: 'Inspect the original cause, plugin options, and render output; return supported docx elements.',
    why(params: { pluginName: string }) {
      return `Plugin render failed: ${params.pluginName}`
    },
  },
  DOCXKIT_PLUGIN_VERSION_MISMATCH: {
    fix: 'Use a plugin version whose docxKit range includes the installed docx-kit version.',
    why(params: { kitVersion: string; pluginName: string; range: string }) {
      return `Plugin "${params.pluginName}" requires docx-kit range "${params.range}" but ${params.kitVersion} is installed`
    },
  },
  DOCXKIT_PREVIEW_FETCH_FAILED: {
    fix: 'Check the DOCX URL, HTTP response, network access, and CORS policy, or pass local bytes.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_PREVIEW_INPUT_INVALID: {
    fix: 'Pass a DOCX Blob, ArrayBuffer, Uint8Array, or URL accepted by the selected renderer.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_PREVIEW_RENDER_FAILED: {
    fix: 'Verify that the input is a valid DOCX archive and inspect the original renderer cause.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_STYLE_CIRCULAR_EXTENDS: {
    fix: 'Remove the circular extends reference so style inheritance forms an acyclic graph.',
    why(params: { chain: string[]; className: string }) {
      return `Circular extends detected for style class "${params.className}". Chain: ${params.chain.join(' → ')}`
    },
  },
  DOCXKIT_STYLE_UNKNOWN_CLASS: {
    fix: 'Define the referenced stylesheet key or correct its spelling in className or extends.',
    why(params: { className: string; extendingClass?: string }) {
      return params.extendingClass === undefined
        ? `Style class not found: "${params.className}"`
        : `Style class "${params.extendingClass}" extends unknown class "${params.className}"`
    },
  },
  DOCXKIT_TABLE_INVALID_COLUMNS: {
    fix: 'Provide table columns and ensure each cell span fits the grid without overlaps.',
    why: (params: { detail: string }) => params.detail,
  },
  DOCXKIT_UNKNOWN_NODE_TYPE: {
    fix: 'Use a supported DSL node type; custom plugins use a plugin node with a registered runtime name.',
    why(params: { nodeType: string }) {
      return `Unknown node type: ${params.nodeType}`
    },
  },
} as const

export type DiagnosticCode = keyof typeof diagnosticDefinitions
/**
 * Existing semantic codes remain stable for error.code comparisons.
 */
export type ErrorCode = LegacyCode<DiagnosticCode>
type LegacyCode<Code> = Code extends `DOCXKIT_${infer Legacy}` ? Legacy : never

// Object.fromEntries loses key/value correlation; this mapped type restores
// the exact relationship guaranteed by the fixed DOCXKIT_ prefix above.
export const ERROR_CODES = Object.fromEntries(
  Object.keys(diagnosticDefinitions).map(code => {
    const legacy = code.slice('DOCXKIT_'.length)
    return [legacy, legacy]
  }),
) as { readonly [Code in ErrorCode]: Code }

/**
 * Full diagnostic codes used in logs and documentation.
 */
export const DIAGNOSTIC_CODES = Object.fromEntries(
  Object.keys(ERROR_CODES).map(code => [code, `DOCXKIT_${code}`]),
) as { readonly [Code in ErrorCode]: `DOCXKIT_${Code}` }

export const diagnostics = defineDiagnostics({
  codes: diagnosticDefinitions,
  reporters: [],
  docsBase: code => `https://docx-kit.ntnyq.dev/errors/${code.toLowerCase()}/`,
})

/**
 * Parameters inferred from the catalog message, plus optional source context.
 * Public declarations do not depend on Nostics implementation types.
 */
export type ErrorParams<Code extends ErrorCode> =
  ((typeof diagnosticDefinitions)[`DOCXKIT_${Code}`] extends {
    why: (params: infer Params) => string
  }
    ? Params
    : object) & {
    cause?: unknown
    sources?: string[]
  }
