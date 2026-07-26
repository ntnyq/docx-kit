/**
 * Content node DSL types.
 *
 * Every piece of content is a **node**. Block nodes represent
 * top-level document sections (headings, paragraphs, tables, images,
 * page breaks, plugins). Inline nodes appear inside text flow.
 *
 * @module dsl/nodes
 */

import type { SectionConfig } from '../document'
import type { BorderRule, DocxStyleRule, StyleSheet } from '../style'
import type { UnitValue } from '../utility'

/**
 * Common properties shared by all nodes.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface BaseNode<TStyles extends StyleSheet = StyleSheet> {
  /**
   * CSS-like class name(s) referencing stylesheet entries.
   *
   * Can be a single string, an array, or a space-separated string.
   */
  className?: string | ClassName<TStyles> | ClassName<TStyles>[]
  /** Optional unique identifier (for templating / references). */
  id?: string
  /** Inline style override for this specific node. */
  style?: DocxStyleRule
}

/**
 * Union of all top-level content node types.
 *
 * @template TStyles — The user's stylesheet type
 */
export type BlockNode<TStyles extends StyleSheet = StyleSheet> =
  | BookmarkNode<TStyles>
  | BulletListNode<TStyles>
  | CheckboxNode<TStyles>
  | ColumnBreakNode
  | CommentNode<TStyles>
  | FootnoteNode<TStyles>
  | HeadingNode<TStyles>
  | HyperlinkNode<TStyles>
  | ImageNode<TStyles>
  | MathNode
  | NumberedListNode<TStyles>
  | PageBreakNode
  | ParagraphNode<TStyles>
  | PluginNode<string, unknown, TStyles>
  | RevisionNode<TStyles>
  | SectionBreakNode
  | TableNode<Record<string, unknown>, TStyles>
  | TextBoxNode<TStyles>
  | ThematicBreakNode<TStyles>

// ---- Utility types ----

/**
 * A named bookmark that can be targeted by an internal hyperlink.
 */
export interface BookmarkNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Text or styled text runs contained by the bookmark. */
  children: (string | TextNode<TStyles>)[]
  /** Stable bookmark name. */
  name: string
  type: 'bookmark'
}

// ---- Block nodes ----

/**
 * A structured list item (for rich bullet/numbered items).
 *
 * @template TStyles — The user's stylesheet type
 */
export interface BulletItem<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Optional inline children override (instead of text). */
  children?: InlineNode<TStyles>[]
  /** Per-item nested list level (0–8). Overrides the list-level default. */
  level?: number
  /** Item text content, used when `children` is absent. */
  text?: string
}

/**
 * A bullet list node.
 *
 * Supports plain string items or structured items with children.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface BulletListNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** List items — strings or structured items. */
  items: (string | BulletItem<TStyles>)[]
  type: 'bulletList'
  /** Bullet character / style. Default: `'•'`. */
  bullet?: string
  /** Nested list level (0 = top-level). */
  level?: number
}

/**
 * A Word checkbox content control.
 */
export interface CheckboxNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  type: 'checkbox'
  /** Accessible alias for the checkbox control. */
  alias?: string
  /** Whether the checkbox is checked. */
  checked?: boolean
  /** Checked-state glyph and font. */
  checkedState?: CheckboxSymbol
  /** Optional label rendered after the checkbox. */
  label?: string
  /** Unchecked-state glyph and font. */
  uncheckedState?: CheckboxSymbol
}

/** Checkbox glyph customization. */
export interface CheckboxSymbol {
  /** Font containing the glyph. */
  font?: string
  /** Unicode code point expressed as a hexadecimal string. */
  value?: string
}

/**
 * Extract valid class name keys from a stylesheet type.
 *
 * @template TStyles — The user's stylesheet type
 */
export type ClassName<TStyles extends StyleSheet> = Extract<
  keyof TStyles,
  string
>

/**
 * A forced column break.
 */
export interface ColumnBreakNode {
  type: 'columnBreak'
}

/**
 * An annotated content range backed by a Word comment.
 */
export interface CommentNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Comment author. */
  author: string
  /** Annotated inline content. */
  children: InlineNode<TStyles>[]
  /** Comment body paragraphs or plain text. */
  comment: (string | ParagraphNode<TStyles>)[]
  type: 'comment'
  /** ISO-8601 creation timestamp. */
  date?: string
  /** Author initials. */
  initials?: string
}

/**
 * A footnote reference and its document-level footnote content.
 */
export interface FootnoteNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Footnote body paragraphs or plain text. */
  content: (string | ParagraphNode<TStyles>)[]
  type: 'footnote'
}

// ---- Inline nodes ----

/**
 * A heading node (h1–h6).
 *
 * @template TStyles — The user's stylesheet type
 */
export interface HeadingNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Heading level (1 = largest, 6 = smallest). */
  level: 1 | 2 | 3 | 4 | 5 | 6
  /** Heading text content. */
  text: string
  type: 'heading'
}

/**
 * A hyperlink node (inline content that creates a clickable link).
 *
 * @template TStyles — The user's stylesheet type
 */
export interface HyperlinkNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Display text or inline children. */
  children: (string | TextNode<TStyles>)[]
  type: 'hyperlink'
  /** Internal bookmark target (used instead of `url`). */
  anchor?: string
  /** External link target URL. */
  url?: string
}

/**
 * An image node.
 *
 * Supports raw bytes (`Uint8Array`, `ArrayBuffer`, `Blob`) or a file path string.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface ImageNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Image data as bytes, blob, or file path. */
  data: string | ArrayBuffer | Blob | Uint8Array
  type: 'image'
  /** Alt text for accessibility. */
  alt?: string
  /** Display height. */
  height?: UnitValue
  /** Image format hint. Auto-detected if omitted. */
  imageType?: 'bmp' | 'gif' | 'jpeg' | 'jpg' | 'png'
  /** Display width. */
  width?: UnitValue
  /**
   * Floating layout configuration.
   *
   * - `true` enables default floating
   * - Object allows wrap mode and position offsets
   */
  floating?:
    | boolean
    | {
        /** Text wrap mode. */
        wrap?: 'square' | 'tight' | 'topAndBottom'
        /** Horizontal offset. */
        x?: UnitValue
        /** Vertical offset. */
        y?: UnitValue
      }
}

/**
 * Union of inline content node types (appear inside paragraphs).
 *
 * @template TStyles — The user's stylesheet type
 */
export type InlineNode<TStyles extends StyleSheet = StyleSheet> =
  | BookmarkNode<TStyles>
  | CheckboxNode<TStyles>
  | CommentNode<TStyles>
  | FootnoteNode<TStyles>
  | HyperlinkNode<TStyles>
  | ImageNode<TStyles>
  | MathNode
  | RevisionNode<TStyles>
  | TextNode<TStyles>

/**
 * A recursive expression used by {@link MathNode}.
 */
export type MathExpression =
  | MathFractionExpression
  | MathFunctionExpression
  | MathIntegralExpression
  | MathRadicalExpression
  | MathScriptExpression
  | MathSumExpression
  | MathTextExpression

/** A mathematical fraction. */
export interface MathFractionExpression {
  denominator: MathExpression[]
  numerator: MathExpression[]
  type: 'fraction'
}

/** A named mathematical function. */
export interface MathFunctionExpression {
  arguments: MathExpression[]
  name: MathExpression[]
  type: 'function'
}

/** A mathematical integral. */
export interface MathIntegralExpression {
  children: MathExpression[]
  type: 'integral'
  subScript?: MathExpression[]
  superScript?: MathExpression[]
}

/**
 * An Office Math (OMML) expression.
 */
export interface MathNode {
  /** Structured mathematical expression tree. */
  children: MathExpression[]
  type: 'math'
}

/** A mathematical radical with an optional degree. */
export interface MathRadicalExpression {
  children: MathExpression[]
  type: 'radical'
  degree?: MathExpression[]
}

/** Subscript, superscript, or combined script. */
export interface MathScriptExpression {
  children: MathExpression[]
  type: 'script'
  subScript?: MathExpression[]
  superScript?: MathExpression[]
}

/** A mathematical sum. */
export interface MathSumExpression {
  children: MathExpression[]
  type: 'sum'
  subScript?: MathExpression[]
  superScript?: MathExpression[]
}

/** Plain text inside a mathematical expression. */
export interface MathTextExpression {
  text: string
  type: 'text'
}

/**
 * A numbered / ordered list node.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface NumberedListNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** List items. */
  items: (string | BulletItem<TStyles>)[]
  type: 'numberedList'
  /** Nested list level (0 = top-level). */
  level?: number
  /** Starting number (default: 1). */
  start?: number
  /**
   * Numbering format.
   *
   * {@link `LevelFormat.DECIMAL`} by default.
   * e.g. `'decimal'`, `'upperRoman'`, `'lowerLetter'`
   */
  numberingFormat?:
    'decimal' | 'lowerLetter' | 'lowerRoman' | 'upperLetter' | 'upperRoman'
}

/**
 * A forced page break.
 */
export interface PageBreakNode {
  type: 'pageBreak'
}

/**
 * A paragraph containing text and/or inline children.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface ParagraphNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  type: 'paragraph'
  /** Inline children (text runs, inline images, etc.). */
  children?: InlineNode<TStyles>[]
  /** Plain-text content (used when `children` is not provided). */
  text?: string
}

/**
 * A plugin-invocation node.
 *
 * Plugins extend the DSL with arbitrary content types.
 *
 * @template TName — Plugin name (string literal)
 * @template TOptions — Plugin-specific options shape
 * @template TStyles — The user's stylesheet type
 */
export interface PluginNode<
  TName extends string = string,
  TOptions = unknown,
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Registered plugin name. */
  name: TName
  /** Plugin-specific options. */
  options: TOptions
  type: 'plugin'
}

/**
 * An inserted or deleted text revision.
 */
export interface RevisionNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Revision author. */
  author: string
  /** Text or styled text runs in the revision. */
  children: (string | TextNode<TStyles>)[]
  /** ISO-8601 revision timestamp. */
  date: string
  /** Document-unique revision identifier. */
  revisionId: number
  /** Revision kind. */
  type: 'deletedText' | 'insertedText'
}

/**
 * Internal marker node that represents a section boundary.
 *
 * Created automatically by {@link DocxBuilder.section} — users should
 * not create this node directly.
 *
 * When the compiler encounters this marker, it starts a new section
 * with the provided configuration.
 */
export interface SectionBreakNode {
  type: 'sectionBreak'
  /** Optional per-section page/header/footer overrides. */
  config?: SectionConfig
}

/** Table-level border configuration. */
export interface TableBordersConfig {
  bottom?: BorderRule
  insideHorizontal?: BorderRule
  insideVertical?: BorderRule
  left?: BorderRule
  right?: BorderRule
  top?: BorderRule
}

/** Per-cell style resolver. */
export type TableCellStyleResolver<
  TData extends Record<string, unknown> = Record<string, unknown>,
> = (
  value: TData[keyof TData],
  row: TData,
  rowIndex: number,
  column: TableColumn<TData>,
) => DocxStyleRule

/**
 * A column definition for a table.
 *
 * @template TData — The row data type
 */
export interface TableColumn<
  TData extends Record<string, unknown> = Record<string, unknown>,
> {
  /** Key in the data object this column maps to. */
  key: Extract<keyof TData, string>
  /** Column header text. */
  title: string
  /** Cell text alignment. */
  align?: 'center' | 'left' | 'right'
  /** Per-column data cell style or resolver. */
  cellStyle?: DocxStyleRule | TableCellStyleResolver<TData>
  /**
   * Span multiple columns horizontally.
   *
   * Applied to all cells in this column.
   */
  colSpan?: number
  /** Per-column header cell style. */
  headerCellStyle?: DocxStyleRule
  /**
   * Span multiple rows vertically (per-cell via data hints).
   *
   * Set a keyed `_${key}_rowSpan: N` hint on an individual data object,
   * use `_rowSpan: N` for every cell in that row, or keep this as a static
   * column default.
   */
  rowSpan?: number
  /** Column width. Supports percentage strings (e.g. `"30%"`). */
  width?: UnitValue
  /**
   * Custom cell renderer.
   *
   * Receives the raw value, full row data, and row index.
   * Returns a string or inline nodes.
   */
  render?: (
    value: TData[keyof TData],
    row: TData,
    index: number,
  ) => string | InlineNode[]
}

/** Floating table positioning. */
export interface TableFloatingOptions {
  /** Bottom distance from surrounding text. */
  bottomFromText?: UnitValue
  /** Horizontal anchor. */
  horizontalAnchor?: 'margin' | 'page' | 'text'
  /** Left distance from surrounding text. */
  leftFromText?: UnitValue
  /** Whether the table may overlap other floating objects. */
  overlap?: boolean
  /** Right distance from surrounding text. */
  rightFromText?: UnitValue
  /** Top distance from surrounding text. */
  topFromText?: UnitValue
  /** Vertical anchor. */
  verticalAnchor?: 'margin' | 'page' | 'text'
  /** Absolute horizontal offset. */
  x?: UnitValue
  /** Absolute vertical offset. */
  y?: UnitValue
  /** Relative horizontal placement. */
  relativeHorizontalPosition?:
    'center' | 'inside' | 'left' | 'outside' | 'right'
  /** Relative vertical placement. */
  relativeVerticalPosition?:
    'bottom' | 'center' | 'inline' | 'inside' | 'outside' | 'top'
}

/** Native Word table-look flags. */
export interface TableLookOptions {
  firstColumn?: boolean
  firstRow?: boolean
  lastColumn?: boolean
  lastRow?: boolean
  noHBand?: boolean
  noVBand?: boolean
}

/**
 * A table node with column definitions and data rows.
 *
 * @template TData — The row data type
 * @template TStyles — The user's stylesheet type
 */
export interface TableNode<
  TData extends Record<string, unknown> = Record<string, unknown>,
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Column definitions. */
  columns: TableColumn<TData>[]
  /** Row data objects. */
  data: TData[]
  type: 'table'
  /** Horizontal table alignment. */
  alignment?: 'center' | 'left' | 'right'
  /** Show table borders. */
  bordered?: boolean
  /** Explicit outer and inner table borders. */
  borders?: TableBordersConfig
  /** Default data cell style or per-cell resolver. */
  cellStyle?: DocxStyleRule | TableCellStyleResolver<TData>
  /** Floating table positioning. Enables side-by-side layouts. */
  floating?: TableFloatingOptions
  /** Show header row (default: `true`). */
  header?: boolean
  /** Style for header cells. */
  headerCellStyle?: DocxStyleRule
  /** Word table layout algorithm. */
  layout?: 'autofit' | 'fixed'
  /** Alternate row shading. */
  striped?: boolean
  /** Native Word table style ID. */
  styleName?: string
  /** Native Word table-look flags. */
  tableLook?: TableLookOptions
  /** Render the table grid from right to left. */
  visuallyRightToLeft?: boolean
  /** Overall table width. */
  width?: UnitValue
}

/**
 * A Word text box positioned as a block-level DrawingML shape.
 */
export interface TextBoxNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Text box dimensions and positioning. */
  box: TextBoxOptions
  type: 'textBox'
  /** Inline children (used instead of `text`). */
  children?: InlineNode<TStyles>[]
  /** Plain text content. */
  text?: string
}

/** Text box dimensions and positioning. */
export interface TextBoxOptions {
  /** Text box width. */
  width: UnitValue
  /** Text box height. */
  height?: UnitValue
  /** Horizontal offset. */
  left?: UnitValue
  /** Shape positioning mode. */
  position?: 'absolute' | 'relative' | 'static'
  /** Vertical offset. */
  top?: UnitValue
  /** Text wrapping behavior. */
  wrap?: 'none' | 'square'
}

/**
 * A text run node (inline content).
 *
 * @template TStyles — The user's stylesheet type
 */
export interface TextNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Text content. */
  text: string
  type: 'text'
}

/**
 * A horizontal thematic break rendered as a paragraph border.
 */
export interface ThematicBreakNode<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  type: 'thematicBreak'
}
