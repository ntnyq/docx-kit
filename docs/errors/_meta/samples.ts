import type { ErrorSamples } from './types'

export const samples: ErrorSamples = {
  HEADER_FOOTER_INVALID_CONTENT: {},
  EXPORT_FAILED: {
    format: 'blob',
  },
  IMAGE_INVALID_DATA: {
    detail: 'Image data is empty or null',
  },
  MANIFEST_INVALID: {
    detail:
      'Plugin manifest field "main" must be a safe relative path beginning with "./"',
  },
  MANIFEST_MISSING: {
    detail: 'Plugin manifest not found for package "example-plugin"',
  },
  MICROSOFT_URL_REQUIRED: {
    detail: 'Microsoft renderer requires a URL string',
  },
  PLUGIN_DEPENDENCY_MISSING: {
    dependency: 'example-dependency',
  },
  PLUGIN_LOAD_FAILED: {
    detail: 'Plugin execution blocked by security policy: example',
    reason: 'security',
  },
  PLUGIN_NOT_REGISTERED: {
    pluginName: 'qrcode',
  },
  PLUGIN_RENDER_FAILED: {
    pluginName: 'example',
  },
  PLUGIN_VERSION_MISMATCH: {
    kitVersion: '0.7.0',
    pluginName: 'example',
    range: '^0.6.0',
  },
  PREVIEW_FETCH_FAILED: {
    detail: 'Failed to fetch DOCX: HTTP 404',
  },
  PREVIEW_INPUT_INVALID: {
    detail: 'Unsupported preview input type',
  },
  PREVIEW_RENDER_FAILED: {
    detail: 'Failed to render DOCX preview',
  },
  STYLE_CIRCULAR_EXTENDS: {
    chain: ['body', 'title', 'body'],
    className: 'body',
  },
  STYLE_UNKNOWN_CLASS: {
    className: 'boddy',
  },
  TABLE_INVALID_COLUMNS: {
    detail: 'Table must have at least one column',
  },
  UNKNOWN_NODE_TYPE: {
    nodeType: 'unknown',
  },
} satisfies ErrorSamples
