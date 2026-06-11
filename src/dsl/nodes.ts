/**
 * Content node DSL types.
 *
 * Every piece of content is a **node**. Block nodes represent
 * top-level document sections (headings, paragraphs, tables, images,
 * page breaks, plugins). Inline nodes appear inside text flow.
 *
 * @module dsl/nodes
 */

import type { SectionConfig } from '../types/document'
import type { DocxStyleRule, StyleSheet } from '../types/style'
import type { UnitValue } from '../types/utility'

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
  | BulletListNode<TStyles>
  | HeadingNode<TStyles>
  | HyperlinkNode<TStyles>
  | ImageNode<TStyles>
  | NumberedListNode<TStyles>
  | PageBreakNode
  | ParagraphNode<TStyles>
  | PluginNode<string, unknown, TStyles>
  | SectionBreakNode
  | TableNode<Record<string, unknown>, TStyles>

// ---- Utility types ----

/**
 * A structured list item (for rich bullet/numbered items).
 *
 * @template TStyles — The user's stylesheet type
 */
export interface BulletItem<
  TStyles extends StyleSheet = StyleSheet,
> extends BaseNode<TStyles> {
  /** Item text content. */
  text: string
  /** Optional inline children override (instead of text). */
  children?: InlineNode<TStyles>[]
}

// ---- Block nodes ----

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
 * Extract valid class name keys from a stylesheet type.
 *
 * @template TStyles — The user's stylesheet type
 */
export type ClassName<TStyles extends StyleSheet> = Extract<
  keyof TStyles,
  string
>

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
  /** Link target URL. */
  url: string
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
  | HyperlinkNode<TStyles>
  | ImageNode<TStyles>
  | TextNode<TStyles>

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
    | 'decimal'
    | 'lowerLetter'
    | 'lowerRoman'
    | 'upperLetter'
    | 'upperRoman'
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
  /**
   * Span multiple columns horizontally.
   *
   * Applied to all cells in this column.
   */
  colSpan?: number
  /**
   * Span multiple rows vertically (per-cell via data hints).
   *
   * Set `rowSpan` on individual data objects using `_rowSpan: N` or keep it
   * as a static column default.
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
  /** Show table borders. */
  bordered?: boolean
  /** Default cell style for data rows. */
  cellStyle?: DocxStyleRule
  /** Show header row (default: `true`). */
  header?: boolean
  /** Style for header cells. */
  headerCellStyle?: DocxStyleRule
  /** Alternate row shading. */
  striped?: boolean
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
