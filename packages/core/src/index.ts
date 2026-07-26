/**
 * @docxkit/core — CSS-like DOCX generation core engine.
 *
 * Built on `dolanmiu/docx`, this package provides the builder API,
 * compiler pipeline, style resolution, DSL types, and cross-platform utilities.
 *
 * @packageDocumentation
 */

// ---------- Cross-platform utilities ----------
export { createImageRun } from './utils/image'
// ---------- Builder ----------
export { DocxBuilder } from './builder/DocxBuilder'

export { validateManifest } from './loader/manifest'

export { dataUrlToUint8Array } from './utils/dataUrl'
export { createDocx, renderDocx } from './builder/createDocx'

// ---------- Loader ----------
export { createPluginLoader, PluginLoader } from './loader/PluginLoader'
// ---------- Re-exports from @docxkit/types ----------
// Value exports
export {
  definePlugin,
  defineStyles,
  DocxKitError,
  ERROR_CODES,
  inlineImg,
  span,
  type ErrorCode,
} from '@docxkit/types'
export type { PluginManifest } from './loader/manifest'

// ---------- Compiler internals (for plugin authors) ----------
export type { CompilationSession } from './compiler/numbers'

// ---------- Local types ----------
export type { DocxSchema, RenderDocxOptions } from './builder/createDocx'

export type {
  PluginLoaderOptions,
  PluginLoadResult,
  PluginManifestAuthorizer,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'

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
