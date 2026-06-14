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

// ---------- Document types ----------
export type {
  DocxKitConfig,
  DocxPreset,
  DocxTheme,
  HeaderFooterConfig,
  HeaderFooterContent,
  Orientation,
  PageConfig,
  PageSize,
  SectionConfig,
  ThemeColors,
  ThemeFonts,
  ThemeSpacing,
} from './document'

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
  TextAlign,
  TextStyleRule,
  VerticalAlign,
} from './style'

// ---------- DSL node types ----------
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
