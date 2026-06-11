/**
 * docx-kit Shared API — re-exported by both `docx-kit` (browser) and `docx-kit/node`.
 *
 * @module docx-kit/shared
 * @packageDocumentation
 */

// ---------- Style ----------
export { defineStyles } from './types/style'

// ---------- Plugin ----------
export { definePlugin } from './types/plugin'

export { qrcodePlugin } from './plugins/qrcode'
// ---------- Built-in plugins ----------
export { calloutPlugin } from './plugins/callout'

export { echartsPlugin } from './plugins/echarts'

// ---------- Builder ----------
export { DocxBuilder } from './builder/DocxBuilder'

export { timelinePlugin } from './plugins/timeline'
// ---------- Errors ----------
export { DocxKitError, ERROR_CODES } from './errors'
export { watermarkPlugin } from './plugins/watermark'
// ---------- Cross-platform utilities ----------
export { dataUrlToUint8Array } from './utils/dataUrl'
export { codeBlockPlugin } from './plugins/code-block'
export { coverPagePlugin } from './plugins/cover-page'
export { dataTablePlugin } from './plugins/data-table'
export { pageNumberPlugin } from './plugins/page-number'
export { createDocx, renderDocx } from './builder/createDocx'
export { propertyTablePlugin } from './plugins/property-table'
export { meetingMinutesPlugin } from './plugins/meeting-minutes'
export { signatureBlockPlugin } from './plugins/signature-block'

// ---------- Presets ----------
export {
  academicPreset,
  classicPreset,
  modernPreset,
  PRESET_LIST,
  usePreset,
} from './presets'

// ---------- Types ----------
export type { ErrorCode } from './errors'
export type { DocxPreset } from './presets'
export type { DocxSchema } from './builder/createDocx'

// Plugin option types
export type { CalloutOptions } from './plugins/callout'

export type { QRCodePluginOptions } from './plugins/qrcode'

export type { WatermarkOptions } from './plugins/watermark'
export type { CodeBlockOptions } from './plugins/code-block'
export type { CoverPageOptions } from './plugins/cover-page'
export type { EChartsPluginOptions } from './plugins/echarts'
export type { PageNumberOptions } from './plugins/page-number'
export type { TimelineEvent, TimelineOptions } from './plugins/timeline'
export type {
  AgendaItem,
  MeetingMinutesOptions,
} from './plugins/meeting-minutes'
export type {
  PropertyItem,
  PropertyTableOptions,
} from './plugins/property-table'
export type {
  ColAlign,
  ColFormat,
  DataTableOptions,
} from './plugins/data-table'
export type {
  SignatureBlockOptions,
  SignatureParty,
} from './plugins/signature-block'
// Plugin system types
export type {
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './types/plugin'
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

// Builder types
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
