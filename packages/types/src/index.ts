/**
 * @docxkit/types — Shared type definitions for docx-kit monorepo.
 *
 * This package exports all shared types, utility types, and
 * type-level helper functions used across @docxkit/* packages.
 *
 * @packageDocumentation
 */

export { defineStyles } from './style'

export { definePlugin } from './plugin'

// ---------- DSL helpers (runtime) ----------
export { inlineImg, span } from './dsl/helpers'

// ---------- Errors (runtime) ----------
export { DocxKitError, ERROR_CODES, type ErrorCode } from './errors'

// ---------- Plugin types ----------
export type {
  BuiltinPluginMap,
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './plugin'

// ---------- Utility types ----------
export type {
  Dict,
  HexColor,
  LiteralUnion,
  MaybePromise,
  StyleToken,
  StyleTokenCategory,
  ThemeToken,
  UnitValue,
} from './utility'

// ---------- Style types ----------
export type {
  BorderRule,
  BorderStyle,
  CellStyleRule,
  DocxStyleRule,
  FontWeight,
  HighlightColor,
  ParagraphStyleRule,
  StyleSheet,
  StyleSheetEntry,
  TabStopRule,
  TextAlign,
  TextStyleRule,
  VerticalAlign,
} from './style'

// ---------- Document types ----------
export type {
  DocumentMetadata,
  DocxKitConfig,
  DocxPreset,
  DocxTheme,
  EmbeddedFont,
  HeaderFooterConfig,
  HeaderFooterContent,
  Orientation,
  PageBorderConfig,
  PageConfig,
  PageNumberConfig,
  PageSize,
  SectionColumn,
  SectionColumnsConfig,
  SectionConfig,
  SectionLineNumberConfig,
  SectionType,
  ThemeColors,
  ThemeFonts,
  ThemeSpacing,
} from './document'

// ---------- DSL node types ----------
export type {
  BaseNode,
  BlockNode,
  BookmarkNode,
  BulletItem,
  BulletListNode,
  CheckboxNode,
  CheckboxSymbol,
  ClassName,
  ColumnBreakNode,
  CommentNode,
  FootnoteNode,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  InlineNode,
  MathExpression,
  MathFractionExpression,
  MathFunctionExpression,
  MathIntegralExpression,
  MathNode,
  MathRadicalExpression,
  MathScriptExpression,
  MathSumExpression,
  MathTextExpression,
  NumberedListNode,
  PageBreakNode,
  ParagraphNode,
  PluginNode,
  RevisionNode,
  SectionBreakNode,
  TableBordersConfig,
  TableCellStyleResolver,
  TableColumn,
  TableFloatingOptions,
  TableLookOptions,
  TableNode,
  TextBoxNode,
  TextBoxOptions,
  TextNode,
  ThematicBreakNode,
} from './dsl/nodes'
