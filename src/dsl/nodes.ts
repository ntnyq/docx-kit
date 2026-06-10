/**
 * Content node DSL types.
 *
 * Every piece of content is a **node**. Block nodes represent
 * top-level document sections (headings, paragraphs, tables, images,
 * page breaks, plugins). Inline nodes appear inside text flow.
 *
 * @module dsl/nodes
 */

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
  | HeadingNode<TStyles>
  | ImageNode<TStyles>
  | PageBreakNode
  | ParagraphNode<TStyles>
  | PluginNode<string, unknown, TStyles>
  | TableNode<Record<string, unknown>, TStyles>

// ---- Utility types ----

/**
 * Extract valid class name keys from a stylesheet type.
 *
 * @template TStyles — The user's stylesheet type
 */
export type ClassName<TStyles extends StyleSheet> = Extract<
  keyof TStyles,
  string
>

// ---- Block nodes ----

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

// ---- Inline nodes ----

/**
 * Union of inline content node types (appear inside paragraphs).
 *
 * @template TStyles — The user's stylesheet type
 */
export type InlineNode<TStyles extends StyleSheet = StyleSheet> =
  | ImageNode<TStyles>
  | TextNode<TStyles>

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
