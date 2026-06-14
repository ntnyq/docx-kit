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

// ---------- Local types ----------
export type { DocxSchema } from './builder/createDocx'

export type { PluginManifest } from './loader/manifest'

// ---------- Compiler internals (for plugin authors) ----------
export type { CompilationSession } from './compiler/numbers'
export type {
  PluginLoadResult,
  PluginSecurityPolicy,
  PluginSource,
} from './loader/PluginLoader'
// Type exports
export type {
  BaseNode,
  BlockNode,
  BorderRule,
  BorderStyle,
  BulletItem,
  BulletListNode,
  ClassName,
  Dict,
  DocxKitConfig,
  DocxPlugin,
  DocxPreset,
  DocxStyleRule,
  DocxTheme,
  FontWeight,
  HeaderFooterConfig,
  HeaderFooterContent,
  HeadingNode,
  HexColor,
  HighlightColor,
  HyperlinkNode,
  ImageNode,
  InlineNode,
  LiteralUnion,
  MaybePromise,
  NumberedListNode,
  Orientation,
  PageBreakNode,
  PageConfig,
  PageSize,
  ParagraphNode,
  PluginNode,
  PluginRegistry,
  PluginRenderContext,
  SectionBreakNode,
  SectionConfig,
  StyleSheet,
  StyleSheetEntry,
  StyleToken,
  StyleTokenCategory,
  TableColumn,
  TableNode,
  TextAlign,
  TextNode,
  ThemeColors,
  ThemeFonts,
  ThemeSpacing,
  ThemeToken,
  UnitValue,
  VerticalAlign,
} from '@docxkit/types'
