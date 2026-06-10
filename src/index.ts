/**
 * docx-kit — A CSS-like DSL for building Word (.docx) documents.
 *
 * ## Quick start
 *
 * ```ts
 * import { createDocx, defineStyles } from 'docx-kit'
 *
 * const styles = defineStyles({ title: { fontSize: 28, fontWeight: 'bold' } })
 *
 * const doc = createDocx({ styles })
 * doc.h1('Hello', { className: 'title' })
 * doc.p('World')
 * await doc.save('output.docx')
 * ```
 *
 * ## Platform-specific imports
 *
 * For Node.js–only APIs (filesystem write, Buffer-based base64 decode):
 * ```ts
 * import { saveDocument, dataUrlToUint8Array } from 'docx-kit/node'
 * ```
 *
 * For browser-only APIs (atob-based base64 decode, Blob handling):
 * ```ts
 * import { dataUrlToUint8Array, normalizeImageData } from 'docx-kit/browser'
 * ```
 *
 * @module docx-kit
 * @packageDocumentation
 */

// ---------- Style ----------
export { defineStyles } from './types/style'

// ---------- Plugin ----------
export { definePlugin } from './types/plugin'

// ---------- Builder ----------
export { DocxBuilder } from './builder/DocxBuilder'

// ---------- Errors ----------
export { DocxKitError, ERROR_CODES } from './errors'

// ---------- Built-in plugins ----------
export { qrcodePlugin } from './plugins/qrcode/index'

// ---------- Cross-platform utilities ----------
export { dataUrlToUint8Array } from './utils/dataUrl'

export { echartsPlugin } from './plugins/echarts/index'
// ---------- Factory ----------
export { createDocx, renderDocx } from './builder/createDocx'

// ---------- Types ----------
export type { ErrorCode } from './errors'
export type { DocxSchema } from './builder/createDocx'
export type { QRCodePluginOptions } from './plugins/qrcode/index'
export type { EChartsPluginOptions } from './plugins/echarts/index'

// Plugin system types
export type {
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './types/plugin'

// Document config types
export type {
  DocxKitConfig,
  DocxTheme,
  Orientation,
  PageConfig,
  PageSize,
} from './types/document'

// Style types
export type {
  BorderRule,
  BorderStyle,
  DocxStyleRule,
  FontWeight,
  StyleSheet,
  TextAlign,
  VerticalAlign,
} from './types/style'

// DSL node types
export type {
  BaseNode,
  BlockNode,
  ClassName,
  HeadingNode,
  ImageNode,
  InlineNode,
  PageBreakNode,
  ParagraphNode,
  PluginNode,
  TableColumn,
  TableNode,
  TextNode,
} from './dsl/nodes'
