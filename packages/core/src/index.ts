/**
 * @docxkit/core — CSS-like DOCX generation core engine.
 *
 * Built on `dolanmiu/docx`, this package provides the builder API,
 * compiler pipeline, style resolution, DSL types, and cross-platform utilities.
 *
 * @packageDocumentation
 */

// ---------- Style ----------
export { defineStyles } from './types/style'

// ---------- Plugin ----------
export { definePlugin } from './types/plugin'

export { createImageRun } from './utils/image'

// ---------- DSL helpers ----------
export { inlineImg, span } from './dsl/helpers'
// ---------- Builder ----------
export { DocxBuilder } from './builder/DocxBuilder'

// ---------- Errors ----------
export { DocxKitError, ERROR_CODES } from './errors'

export { validateManifest } from './loader/manifest'
// ---------- Cross-platform utilities ----------
export { dataUrlToUint8Array } from './utils/dataUrl'

export { createDocx, renderDocx } from './builder/createDocx'

// ---------- Loader ----------
export { createPluginLoader, PluginLoader } from './loader/PluginLoader'

// ---------- Types (value exports) ----------
export type { ErrorCode } from './errors'

// Builder types
export type { DocxSchema } from './builder/createDocx'
export type { PluginManifest } from './loader/manifest'

// ---------- Compiler internals (for plugin authors) ----------
export type { CompilationSession } from './compiler/numbers'

// Plugin types
export type {
  DocxPlugin,
  PluginRegistry,
  PluginRenderContext,
} from './types/plugin'

export type {
  PluginLoadResult,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'
// Utility types
export type {
  Dict,
  HexColor,
  LiteralUnion,
  MaybePromise,
  StyleToken,
  StyleTokenCategory,
  ThemeToken,
  UnitValue,
} from './types/utility'
// Style types
export type {
  BorderRule,
  BorderStyle,
  DocxStyleRule,
  FontWeight,
  HighlightColor,
  StyleSheet,
  StyleSheetEntry,
  TextAlign,
  VerticalAlign,
} from './types/style'
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
