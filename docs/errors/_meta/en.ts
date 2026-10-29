import type { ErrorNotes } from './types'

export const notes = {
  EXPORT_FAILED: {
    causes: ['Invalid document resources or packaging failure.'],
    title: 'Document export failed',
    steps: [
      'Inspect error.cause locally to identify the failing resource.',
      'Use an export format supported by the current platform and verify referenced resources.',
    ],
    summary:
      'Packaging the compiled document failed. The original cause remains available to the caller.',
  },
  HEADER_FOOTER_INVALID_CONTENT: {
    title: 'Invalid header or footer output',
    causes: [
      'A plugin returned a bare TextRun or another unsupported child element.',
    ],
    steps: [
      'Return Paragraph or Table instances for header and footer content.',
      'Wrap inline runs inside a Paragraph.',
    ],
    summary:
      'A header/footer plugin returned an element other than a Paragraph or Table.',
  },
  IMAGE_INVALID_DATA: {
    causes: ['Empty image data, unsupported format, or an unresolved path.'],
    title: 'Invalid image data',
    steps: [
      'Supply supported non-empty image bytes or a data URL.',
      'For file paths use docx-kit/node; in browsers configure resolveImage.',
      'Check imageType when automatic format detection is unavailable.',
    ],
    summary:
      'Image content is empty, cannot be identified, or requires a path resolver.',
  },
  MANIFEST_INVALID: {
    summary: 'The manifest has an invalid field, version range, or entry path.',
    title: 'Invalid plugin manifest',
    causes: [
      'Missing required fields, malformed version ranges, or paths escaping the package.',
    ],
    steps: [
      'Correct the field named in the message.',
      'Use non-empty strings, valid semver, and safe relative paths beginning with ./.',
      'Keep the entry inside the plugin package.',
    ],
  },
  MANIFEST_MISSING: {
    title: 'Plugin manifest missing',
    causes: [
      'The published package omitted its manifest or the requested path returned an error.',
    ],
    steps: [
      'Include the manifest at the package root.',
      'Verify that the manifest path or URL resolves before importing executable code.',
    ],
    summary:
      'External loading could not locate the required docx-kit.plugin.json.',
  },
  MICROSOFT_URL_REQUIRED: {
    title: 'Microsoft preview requires a URL',
    causes: [
      'Passing local bytes, a Blob URL, or a non-HTTP(S) URL to Microsoft preview.',
    ],
    steps: [
      'Host the document at a publicly accessible HTTP(S) URL.',
      'For local blobs or private files, use the local docx-preview renderer.',
    ],
    summary:
      'Microsoft Office Online preview requires a public HTTP(S) DOCX URL.',
  },
  PLUGIN_DEPENDENCY_MISSING: {
    title: 'Plugin dependency missing',
    causes: [
      'An integration explicitly reports a missing or incompatible dependency.',
    ],
    steps: [
      'If a custom integration raises this code, check its dependency validation.',
      'Install the dependency with a version required by the plugin.',
    ],
    summary:
      'A public code for plugin dependency failures. Built-in loaders currently do not automatically raise this code.',
  },
  PLUGIN_LOAD_FAILED: {
    title: 'Plugin loading failed',
    causes: [
      'An unsupported source type, policy rejection, import error, or invalid plugin export.',
    ],
    steps: [
      'For environment failures, use the appropriate browser or Node.js loader.',
      'For import failures, check the entry, dependencies, and network.',
      'For export/factory failures, return a valid plugin instance and inspect the original cause.',
      'For security failures, review source and permissions against policy; keep untrusted code blocked.',
    ],
    summary:
      'Loading failed because of the platform, security policy, import, module export, or plugin factory.',
  },
  PLUGIN_NOT_REGISTERED: {
    title: 'Plugin not registered',
    causes: [
      'Missing .use(), a factory that was not called, or a runtime name mismatch.',
    ],
    steps: [
      'Call the plugin factory and pass the returned instance to .use() before export.',
      'Match the node name to plugin.name; third-party factory names may differ.',
      'JSON DSL plugins must also be registered.',
    ],
    summary:
      'A plugin node references a runtime name that has not been registered. PDK assertions also reuse this code for invalid plugin instances.',
  },
  PLUGIN_RENDER_FAILED: {
    title: 'Plugin rendering failed',
    causes: [
      'A renderer exception, invalid options, or an unexpected PDK output type.',
    ],
    steps: [
      'Inspect the original cause locally and validate the plugin options.',
      'Return supported docx elements from render.',
      'For PDK assertions verify the expected element types and counts.',
    ],
    summary:
      'The plugin render method threw, or a PDK render assertion rejected its output.',
  },
  PLUGIN_VERSION_MISMATCH: {
    title: 'Incompatible plugin version',
    causes: [
      'An incompatible release pair or an invalid running version/range.',
    ],
    steps: [
      'Compare the range and installed version shown in the message.',
      'Select compatible plugin and docx-kit versions.',
    ],
    summary:
      'The manifest docxKit range does not accept the installed library version.',
  },
  PREVIEW_FETCH_FAILED: {
    summary: 'The DOCX URL could not be fetched successfully.',
    title: 'Preview fetch failed',
    causes: [
      'A non-success HTTP response, network failure, or blocked cross-origin request.',
    ],
    steps: [
      'Check the HTTP response, network, authentication, and CORS settings.',
      'Use local DOCX bytes when remote fetching is unavailable.',
    ],
  },
  PREVIEW_INPUT_INVALID: {
    causes: ['Null input or an unsupported object/string type.'],
    title: 'Invalid preview input',
    steps: [
      'Pass a DOCX Blob, ArrayBuffer, Uint8Array, or URL.',
      'Wait for generation to finish before starting preview.',
    ],
    summary:
      'The preview input is absent or not one of the supported input types.',
  },
  PREVIEW_RENDER_FAILED: {
    causes: ['Corrupt archives, unsupported content, or a renderer exception.'],
    summary: 'The local docx-preview renderer rejected the document.',
    title: 'Preview rendering failed',
    steps: [
      'Check that the input is a valid DOCX ZIP archive.',
      'Inspect the original renderer cause; try opening the same file in Word.',
    ],
  },
  STYLE_CIRCULAR_EXTENDS: {
    causes: ['Self inheritance or a cycle across multiple style classes.'],
    title: 'Circular style inheritance',
    steps: [
      'Follow the chain shown in the message.',
      'Remove a reference that closes the cycle; shared parents are allowed.',
    ],
    summary:
      'The extends chain references a style that is already being resolved.',
  },
  STYLE_UNKNOWN_CLASS: {
    causes: ['A spelling mismatch, missing parent style, or wrong stylesheet.'],
    title: 'Unknown style class',
    steps: [
      'Compare the referenced class with the exact stylesheet keys.',
      'Define missing parent classes before calling defineStyles.',
    ],
    summary:
      'A className or extends reference does not match a stylesheet key.',
  },
  TABLE_INVALID_COLUMNS: {
    title: 'Invalid table grid',
    causes: [
      'Empty columns or cell spans that exceed or overlap the table grid.',
    ],
    steps: [
      'Define at least one column.',
      'Keep rowSpan and colSpan within the grid and remove overlapping cell positions.',
    ],
    summary:
      'The table has no columns or contains invalid, overlapping, or out-of-range spans.',
  },
  UNKNOWN_NODE_TYPE: {
    summary: 'No compiler supports this node in its current context.',
    title: 'Unknown node type',
    causes: [
      'Misspelled node types, block nodes used inline, or nested section breaks.',
    ],
    steps: [
      'Use a documented DSL type.',
      'Represent custom content with a registered plugin node.',
      'Keep sectionBreak nodes at document level, outside nested content.',
    ],
  },
} satisfies ErrorNotes
