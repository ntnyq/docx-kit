/**
 * CSS-like style type definitions for docx-kit.
 *
 * These mirror familiar CSS property names and are compiled
 * into Word OpenXML (docx) constructor options internally.
 *
 * @module style
 */

import { DocxKitError } from './errors'
import type { HexColor, LiteralUnion, StyleToken, UnitValue } from './utility'

/**
 * A single border side descriptor.
 *
 * Follows the CSS `border` shorthand convention (style, width, color).
 */
export interface BorderRule {
  /**
   * Border color (e.g. `"#333"`, named color, or theme token like `"$colors.primary"`).
   */
  color?: HexColor | StyleToken<string>
  /**
   * Line style.
   */
  style?: BorderStyle
  /**
   * Line width (bare number treated as pt).
   */
  width?: UnitValue
}

/**
 * Available border line styles.
 */
export type BorderStyle = 'dashed' | 'dotted' | 'double' | 'none' | 'single'

/**
 * Table-cell level style properties.
 *
 * Used by {@link compileCellStyle} for `TableCell` construction.
 */
export interface CellStyleRule {
  /**
   * Background / shading color (hex, named, or theme token like `"$colors.info"`).
   */
  backgroundColor?: HexColor | StyleToken<string>
  /**
   * Shorthand border for all four sides.
   */
  border?: BorderRule
  /**
   * Bottom border override.
   */
  borderBottom?: BorderRule
  /**
   * Left border override.
   */
  borderLeft?: BorderRule
  /**
   * Right border override.
   */
  borderRight?: BorderRule
  /**
   * Top border override.
   */
  borderTop?: BorderRule
  /**
   * Element height.
   */
  height?: UnitValue
  /**
   * Bottom margin (cell padding).
   */
  marginBottom?: UnitValue
  /**
   * Left margin (cell padding).
   */
  marginLeft?: UnitValue
  /**
   * Right margin (cell padding).
   */
  marginRight?: UnitValue
  /**
   * Top margin (cell padding).
   */
  marginTop?: UnitValue
  /**
   * Vertical alignment (mostly for table cells).
   */
  verticalAlign?: VerticalAlign
  /**
   * Element width.
   */
  width?: UnitValue
  /**
   * CSS-like margin shorthand (used as cell padding).
   */
  margin?:
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
    | UnitValue
}

/**
 * CSS-like style descriptor for a run, paragraph, cell, or table.
 *
 * Combines text, paragraph, and cell-level properties.
 * This is the user-facing style type — individual compiler functions
 * narrow to {@link TextStyleRule}, {@link ParagraphStyleRule}, or
 * {@link CellStyleRule} as appropriate.
 */
export interface DocxStyleRule
  extends CellStyleRule, ParagraphStyleRule, TextStyleRule {}

// ---- Sub-types and enums ----

/**
 * Font weight: keyword `"bold"` / `"normal"`, or numeric 100–900
 * following the CSS `font-weight` spec.
 */
export type FontWeight =
  'bold' | 'normal' | 100 | 200 | 300 | 400 | 500 | 600 | 700 | 800 | 900

/**
 * Text highlight / marker colors (matches Word highlight palette).
 */
export type HighlightColor =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGray'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'green'
  | 'lightGray'
  | 'magenta'
  | 'none'
  | 'red'
  | 'white'
  | 'yellow'

/**
 * Paragraph-level style properties.
 *
 * Used by {@link compileParagraphStyle} for `Paragraph` construction.
 */
export interface ParagraphStyleRule {
  // ---- border ----
  /**
   * Shorthand border for all four sides.
   */
  border?: BorderRule
  /**
   * Bottom border override.
   */
  borderBottom?: BorderRule
  /**
   * Left border override.
   */
  borderLeft?: BorderRule
  /**
   * Right border override.
   */
  borderRight?: BorderRule
  /**
   * Top border override.
   */
  borderTop?: BorderRule

  // ---- paragraph flow ----
  /**
   * Keep lines together on same page.
   */
  keepLines?: boolean
  /**
   * Keep this paragraph with the next one.
   */
  keepNext?: boolean
  /**
   * Line height multiplier or explicit unit value.
   */
  lineHeight?: number | UnitValue
  /**
   * Bottom margin.
   */
  marginBottom?: UnitValue
  /**
   * Left margin.
   */
  marginLeft?: UnitValue
  /**
   * Right margin.
   */
  marginRight?: UnitValue
  /**
   * Top margin.
   */
  marginTop?: UnitValue
  /**
   * Outline level used by Word navigation and table-of-contents fields (0–9).
   */
  outlineLevel?: number
  /**
   * Force page break before this paragraph.
   */
  pageBreakBefore?: boolean
  /**
   * Custom tab stops for this paragraph.
   */
  tabStops?: TabStopRule[]
  /**
   * Horizontal text alignment.
   */
  textAlign?: TextAlign
  /**
   * First-line indent.
   */
  textIndent?: UnitValue
  /**
   * Prevent isolated first or last lines across page boundaries.
   */
  widowControl?: boolean
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
}

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
export type StyleSheet = Record<string, StyleSheetEntry>

/**
 * A single entry in the stylesheet.
 *
 * Extends {@link DocxStyleRule} with an optional `extends` property
 * that allows style classes to inherit from other classes.
 *
 * @example
 * ```ts
 * defineStyles({
 *   baseText: { fontFamily: 'Arial', fontSize: 12, color: '#333' },
 *   heading:  { extends: 'baseText', fontSize: 20, fontWeight: 'bold' },
 *   muted:    { extends: 'baseText', color: '#888' },
 * })
 * ```
 */
export interface StyleSheetEntry extends DocxStyleRule {
  /**
   * Inherit style properties from one or more other style classes.
   */
  extends?: string | string[]
}

/**
 * A paragraph tab stop.
 */
export interface TabStopRule {
  /**
   * Position from the paragraph's left edge.
   */
  position: UnitValue
  /**
   * Optional leader characters before the tab stop.
   */
  leader?: 'dot' | 'hyphen' | 'middleDot' | 'none' | 'underscore'
  /**
   * Tab alignment.
   */
  type:
    | 'bar'
    | 'center'
    | 'clear'
    | 'decimal'
    | 'end'
    | 'left'
    | 'num'
    | 'right'
    | 'start'
}

/**
 * Horizontal text alignment.
 */
export type TextAlign = 'center' | 'justify' | 'left' | 'right'

/**
 * Text-level style properties (font, color, size, weight, etc.).
 *
 * Used by {@link compileTextStyle} for `TextRun` construction.
 *
 * @remarks
 * Convenience boolean shortcuts `bold` and `italic` are provided
 * alongside the canonical `fontWeight` / `fontStyle` properties.
 * When `bold: true` is set, it is equivalent to `fontWeight: 'bold'`.
 * When `italic: true` is set, it is equivalent to `fontStyle: 'italic'`.
 */
export interface TextStyleRule {
  /**
   * Force text to uppercase (small caps-like).
   */
  allCaps?: boolean
  /**
   * Background / shading color (hex, named, or theme token like `"$colors.info"`).
   */
  backgroundColor?: HexColor | StyleToken<string>
  /**
   * Convenience boolean: `true` maps to `fontWeight: 'bold'`.
   * Takes precedence over `fontWeight` when both are set.
   */
  bold?: boolean
  /**
   * Raw OOXML character spacing in twips. Prefer `letterSpacing` for CSS units.
   */
  characterSpacing?: number
  /**
   * Text / foreground color (hex, named, or theme token like `"$colors.primary"`).
   */
  color?: HexColor | StyleToken<string>
  /**
   * Direct passthrough to the underlying `docx` library constructor options.
   * Use for properties not yet covered by the CSS-like mapping.
   */
  docx?: Record<string, unknown>
  /**
   * Double strikethrough toggle.
   */
  doubleStrike?: boolean
  /**
   * Embossed text effect.
   */
  emboss?: boolean
  /**
   * Font family name (e.g. `"Arial"`, or theme token like `"$fonts.heading"`).
   */
  fontFamily?: StyleToken<LiteralUnion<'Arial' | 'Calibri' | 'Times New Roman'>>
  /**
   * Font size (bare number = pt).
   */
  fontSize?: UnitValue
  /**
   * Italic toggle (canonical form). Use `italic?: boolean` for convenience.
   */
  fontStyle?: 'italic' | 'normal'
  /**
   * Font weight: keyword `"bold"` / `"normal"` or numeric 100–900.
   */
  fontWeight?: FontWeight
  /**
   * Text highlight color (background marker).
   */
  highlight?: HighlightColor
  /**
   * Imprinted (engraved) text effect.
   */
  imprint?: boolean
  /**
   * Convenience boolean: `true` maps to `fontStyle: 'italic'`.
   * Takes precedence over `fontStyle` when both are set.
   */
  italic?: boolean
  /**
   * CSS-like character spacing with unit conversion.
   */
  letterSpacing?: UnitValue
  /**
   * Right-to-left run direction.
   */
  rightToLeft?: boolean
  /**
   * Small caps text variant.
   */
  smallCaps?: boolean
  /**
   * Strikethrough toggle.
   */
  strike?: boolean
  /**
   * Sub-script text.
   */
  subScript?: boolean
  /**
   * Super-script text.
   */
  superScript?: boolean
  /**
   * Underline style.
   */
  underline?: 'double' | 'single' | boolean
}

/**
 * Vertical alignment (for table cells).
 */
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
  // Filter out metadata-only keys like `extends` from the resolved rule.
  // This validation can be extended to check for circular inheritance.
  const styleKeys = new Set(Object.keys(styles))
  for (const [name, entry] of Object.entries(styles)) {
    if (entry.extends) {
      const parents = Array.isArray(entry.extends)
        ? entry.extends
        : [entry.extends]
      for (const parent of parents) {
        if (!styleKeys.has(parent)) {
          throw new DocxKitError(
            'STYLE_UNKNOWN_CLASS',
            `Style class "${name}" extends unknown class "${parent}"`,
          )
        }
      }
    }
  }
  return styles
}
