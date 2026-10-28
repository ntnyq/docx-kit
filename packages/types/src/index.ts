/**
 * `@docxkit/types` — Shared type definitions for docx-kit monorepo.
 *
 * This package exports shared type definitions and utility types only.
 * Runtime helpers live in `@docxkit/core`; errors live in `@docxkit/errors`.
 *
 * @packageDocumentation
 */

// ---------- Plugin types ----------
export type {
  BuiltinPluginMap,
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './plugin.js'

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
} from './utility.js'

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
} from './style.js'

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
} from './document.js'

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
} from './dsl/nodes.js'
