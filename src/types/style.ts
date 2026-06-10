/**
 * CSS-like style type definitions for docx-kit.
 *
 * These mirror familiar CSS property names and are compiled
 * into Word OpenXML (docx) constructor options internally.
 *
 * @module types/style
 */

import type { HexColor, LiteralUnion, UnitValue } from './utility'

/**
 * A single border side descriptor.
 *
 * Follows the CSS `border` shorthand convention (style, width, color).
 */
export interface BorderRule {
  /** Border color (e.g. `"#333"` or named color). */
  color?: string | HexColor
  /** Line style. */
  style?: BorderStyle
  /** Line width (bare number treated as pt). */
  width?: UnitValue
}

/** Available border line styles. */
export type BorderStyle = 'dashed' | 'dotted' | 'double' | 'none' | 'single'

/**
 * CSS-like style descriptor for a run, paragraph, cell, or table.
 *
 * Keys mirror familiar CSS property names. Unknown properties that don't
 * map to Word XML are silently ignored.
 */
export interface DocxStyleRule {
  // ---- text ----
  /** Force text to uppercase (small caps-like). */
  allCaps?: boolean
  /** Background / shading color. */
  backgroundColor?: string | HexColor
  // ---- border ----
  /** Shorthand border for all four sides. */
  border?: BorderRule
  /** Bottom border override. */
  borderBottom?: BorderRule
  /** Left border override. */
  borderLeft?: BorderRule
  /** Right border override. */
  borderRight?: BorderRule
  /** Top border override. */
  borderTop?: BorderRule
  /** Text / foreground color. */
  color?: string | HexColor
  // ---- Word-specific escape hatch ----
  /**
   * Direct passthrough to the underlying `docx` library constructor options.
   * Use for properties not yet covered by the CSS-like mapping.
   */
  docx?: Record<string, unknown>

  /** Font family name. */
  fontFamily?: LiteralUnion<'Arial' | 'Calibri' | 'Times New Roman'>
  /** Font size (bare number = pt). */
  fontSize?: UnitValue
  /** Italic toggle. */
  fontStyle?: 'italic' | 'normal'
  /** Font weight: keyword `"bold"` / `"normal"` or numeric 100–900. */
  fontWeight?: FontWeight

  // ---- layout ----
  /** Element height. */
  height?: UnitValue
  // ---- paragraph ----
  /** Character spacing. */
  letterSpacing?: UnitValue
  /** Line height multiplier or explicit unit value. */
  lineHeight?: number | UnitValue
  /** Bottom margin. */
  marginBottom?: UnitValue
  /** Left margin. */
  marginLeft?: UnitValue
  /** Right margin. */
  marginRight?: UnitValue

  /** Top margin. */
  marginTop?: UnitValue
  /** Strikethrough toggle. */
  strike?: boolean
  /** Horizontal text alignment. */
  textAlign?: TextAlign

  /** First-line indent. */
  textIndent?: UnitValue
  /** Underline style. */
  underline?: 'double' | 'single' | boolean
  /** Vertical alignment (mostly for table cells). */
  verticalAlign?: VerticalAlign
  /** Element width. */
  width?: UnitValue
  // ---- spacing / box ----
  /**
   * CSS-like margin shorthand.
   *
   * Supports 1-value, 2-value, and 4-value string syntax
   * (e.g. `"10pt"`, `"10pt 20pt"`, `"10pt 20pt 30pt 40pt"`).
   */
  margin?:
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
    | UnitValue

  /**
   * CSS-like padding shorthand (same syntax as margin).
   */
  padding?:
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
    | UnitValue
}

/**
 * Font weight: keyword `"bold"` / `"normal"`, or numeric 100–900
 * following the CSS `font-weight` spec.
 */
export type FontWeight =
  | 'bold'
  | 'normal'
  | 100
  | 200
  | 300
  | 400
  | 500
  | 600
  | 700
  | 800
  | 900

/**
 * A map of class name → style rule.
 *
 * Used to define reusable named styles referenced via `className` on nodes.
 *
 * @example
 * ```ts
 * const styles = defineStyles({
 *   red:  { color: '#ff0000' },
 *   big:  { fontSize: 20 },
 * })
 * ```
 */
export type StyleSheet = Record<string, DocxStyleRule>

/** Horizontal text alignment. */
export type TextAlign = 'center' | 'justify' | 'left' | 'right'

/** Vertical alignment (for table cells). */
export type VerticalAlign = 'bottom' | 'middle' | 'top'

/**
 * Define a type-safe stylesheet.
 *
 * `fontSize` defaults to pt when passed as a bare number.
 *
 * @param styles - — The stylesheet object
 * @returns The same stylesheet with `const` type inference
 *
 * @example
 * ```ts
 * const styles = defineStyles({
 *   title: { fontSize: 28, fontWeight: 'bold' },
 *   body:  { fontSize: 12, lineHeight: 1.5 },
 * })
 * ```
 */
export function defineStyles<const T extends StyleSheet>(styles: T): T {
  return styles
}
