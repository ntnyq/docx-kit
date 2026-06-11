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

export { qrcodePlugin } from './plugins/qrcode/index'
// ---------- Cross-platform utilities ----------
export { dataUrlToUint8Array } from './utils/dataUrl'
// ---------- Built-in plugins ----------
export { calloutPlugin } from './plugins/callout/index'
export { echartsPlugin } from './plugins/echarts/index'
export { timelinePlugin } from './plugins/timeline/index'
export { watermarkPlugin } from './plugins/watermark/index'
export { codeBlockPlugin } from './plugins/code-block/index'
export { coverPagePlugin } from './plugins/cover-page/index'
export { dataTablePlugin } from './plugins/data-table/index'
// ---------- Factory ----------
export { createDocx, renderDocx } from './builder/createDocx'
export { pageNumberPlugin } from './plugins/page-number/index'
export { propertyTablePlugin } from './plugins/property-table/index'
export { meetingMinutesPlugin } from './plugins/meeting-minutes/index'

export { signatureBlockPlugin } from './plugins/signature-block/index'

// ---------- Types ----------
export type { ErrorCode } from './errors'
export type { DocxSchema } from './builder/createDocx'

// Plugin option types
export type { CalloutOptions } from './plugins/callout/index'
export type { QRCodePluginOptions } from './plugins/qrcode/index'
export type { WatermarkOptions } from './plugins/watermark/index'
export type { CodeBlockOptions } from './plugins/code-block/index'
export type { CoverPageOptions } from './plugins/cover-page/index'
export type { EChartsPluginOptions } from './plugins/echarts/index'
export type { PageNumberOptions } from './plugins/page-number/index'
export type { TimelineEvent, TimelineOptions } from './plugins/timeline/index'
// Plugin system types
export type {
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './types/plugin'
// Sub-types used in plugin options
export type {
  AgendaItem,
  MeetingMinutesOptions,
} from './plugins/meeting-minutes/index'

export type {
  PropertyItem,
  PropertyTableOptions,
} from './plugins/property-table/index'

export type {
  ColAlign,
  ColFormat,
  DataTableOptions,
} from './plugins/data-table/index'

export type {
  SignatureBlockOptions,
  SignatureParty,
} from './plugins/signature-block/index'

// Style types
export type {
  BorderRule,
  BorderStyle,
  DocxStyleRule,
  FontWeight,
  HighlightColor,
  StyleSheet,
  TextAlign,
  VerticalAlign,
} from './types/style'

// Document config types
export type {
  DocxKitConfig,
  DocxTheme,
  HeaderFooterConfig,
  HeaderFooterContent,
  Orientation,
  PageConfig,
  PageSize,
  SectionConfig,
} from './types/document'

// DSL node types
export type {
  BaseNode,
  BlockNode,
  BulletItem,
  BulletListNode,
  ClassName,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  InlineNode,
  NumberedListNode,
  PageBreakNode,
  ParagraphNode,
  PluginNode,
  SectionBreakNode,
  TableColumn,
  TableNode,
  TextNode,
} from './dsl/nodes'
