/**
 * `@docxkit/core` — CSS-like DOCX generation core engine.
 *
 * Built on `dolanmiu/docx`, this package provides the builder API,
 * compiler pipeline, style resolution, DSL types, and cross-platform utilities.
 *
 * @packageDocumentation
 */

// ---------- Cross-platform utilities ----------
export { createImageRun } from './utils/image'
// ---------- DSL and definition helpers ----------
export { inlineImg, span } from './dsl/helpers'
export { normalizeHexColor } from './utils/color'
// ---------- Builder ----------
export { DocxBuilder } from './builder/DocxBuilder'
export { defineStyles } from './style/defineStyles'

export { validateManifest } from './loader/manifest'

export { definePlugin } from './builder/definePlugin'
export { dataUrlToUint8Array } from './utils/dataUrl'

export { createPluginStyle } from './style/pluginStyle'
export { createDocx, renderDocx } from './builder/createDocx'
// ---------- Loader ----------
export { createPluginLoader, PluginLoader } from './loader/PluginLoader'

// ---------- Shared errors ----------
export {
  createDocxKitError,
  DIAGNOSTIC_CODES,
  DocxKitError,
  ERROR_CODES,
  serializeDocxKitError,
} from '@docxkit/errors'
export type { PluginManifest } from './loader/manifest'

// ---------- Compiler internals (for plugin authors) ----------
export type { CompilationSession } from './compiler/numbers'
export type { PluginStyleAdapter } from './style/pluginStyle'

// ---------- Local types ----------
export type { DocxSchema, RenderDocxOptions } from './builder/createDocx'

export type {
  PluginLoaderOptions,
  PluginLoadResult,
  PluginManifestAuthorizer,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'

export type {
  CreateDocxKitErrorOptions,
  DiagnosticCode,
  DocxKitErrorDetails,
  ErrorCode,
  ErrorParams,
  PluginLoadReason,
  SerializedDocxKitError,
} from '@docxkit/errors'

// Type exports
export type {
  BaseNode,
  BlockNode,
  BookmarkNode,
  BorderRule,
  BorderStyle,
  BulletItem,
  BulletListNode,
  CheckboxNode,
  CheckboxSymbol,
  ClassName,
  ColumnBreakNode,
  CommentNode,
  Dict,
  DocumentMetadata,
  DocxKitConfig,
  DocxPlugin,
  DocxPreset,
  DocxStyleRule,
  DocxTheme,
  EmbeddedFont,
  FontWeight,
  FootnoteNode,
  HeaderFooterConfig,
  HeaderFooterContent,
  HeadingNode,
  HexColor,
  HighlightColor,
  HyperlinkNode,
  ImageNode,
  InlineNode,
  LiteralUnion,
  MathExpression,
  MathFractionExpression,
  MathFunctionExpression,
  MathIntegralExpression,
  MathNode,
  MathRadicalExpression,
  MathScriptExpression,
  MathSumExpression,
  MathTextExpression,
  MaybePromise,
  NumberedListNode,
  Orientation,
  PageBorderConfig,
  PageBreakNode,
  PageConfig,
  PageNumberConfig,
  PageSize,
  ParagraphNode,
  PluginNode,
  PluginRegistry,
  PluginRenderContext,
  RevisionNode,
  SectionBreakNode,
  SectionColumn,
  SectionColumnsConfig,
  SectionConfig,
  SectionLineNumberConfig,
  SectionType,
  StyleSheet,
  StyleSheetEntry,
  StyleToken,
  StyleTokenCategory,
  TableBordersConfig,
  TableCellStyleResolver,
  TableColumn,
  TableFloatingOptions,
  TableLookOptions,
  TableNode,
  TabStopRule,
  TextAlign,
  TextBoxNode,
  TextBoxOptions,
  TextNode,
  ThematicBreakNode,
  ThemeColors,
  ThemeFonts,
  ThemeSpacing,
  ThemeToken,
  UnitValue,
  VerticalAlign,
} from '@docxkit/types'
