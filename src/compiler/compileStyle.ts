/**
 * Style compiler — maps `DocxStyleRule` (CSS-like) → `docx` constructor options.
 *
 * Each exported function converts one aspect of a style rule into the shape
 * expected by the underlying `docx` library.
 *
 * @module compiler/compileStyle
 */

import {
  AlignmentType,
  BorderStyle,
  ShadingType,
  VerticalAlign,
  WidthType,
} from 'docx'
import { parseShorthandTwip, toPtHalf, toTwip } from './units'
import type { DocxStyleRule } from '../types/style'

// ---------- Text / Run ----------

/**
 * Compile all border properties from a `DocxStyleRule`.
 *
 * Returns `undefined` if no borders are specified.
 *
 * @param style - — The resolved style rule
 * @returns Border options for `docx` or `undefined`
 *
 * @example
 * ```ts
 * compileBorder({ border: { style: 'single', color: '#333', width: '1pt' } })
 * // => { top: {...}, right: {...}, bottom: {...}, left: {...} }
 * ```
 */
export function compileBorder(style: DocxStyleRule) {
  const b = style.border
  const bt = style.borderTop ?? b
  const br = style.borderRight ?? b
  const bb = style.borderBottom ?? b
  const bl = style.borderLeft ?? b

  if (!bt && !br && !bb && !bl) {
    return undefined
  }

  return {
    bottom: bb ? compileSingleBorder(bb!) : undefined,
    left: bl ? compileSingleBorder(bl!) : undefined,
    right: br ? compileSingleBorder(br!) : undefined,
    top: bt ? compileSingleBorder(bt!) : undefined,
  }
}

// ---------- Paragraph ----------

/**
 * Compile a `DocxStyleRule` into `docx` table cell options
 * (vertical alignment, margins, shading).
 *
 * @param style - — The resolved style rule
 * @returns Options object suitable for `new TableCell(...)`
 *
 * @example
 * ```ts
 * compileCellStyle({ verticalAlign: 'middle', backgroundColor: '#f0f0f0' })
 * // => { verticalAlign: VerticalAlign.CENTER, shading: { fill: 'f0f0f0', ... } }
 * ```
 */
export function compileCellStyle(style: DocxStyleRule) {
  const shorthand =
    style.margin == null ? undefined : parseShorthandTwip(style.margin)

  return {
    verticalAlign: compileVerticalAlign(style.verticalAlign),
    margins: {
      bottom: toTwip(style.marginBottom) ?? shorthand?.bottom,
      left: toTwip(style.marginLeft) ?? shorthand?.left,
      right: toTwip(style.marginRight) ?? shorthand?.right,
      top: toTwip(style.marginTop) ?? shorthand?.top,
    },
    shading:
      style.backgroundColor == null
        ? undefined
        : {
            fill: normalizeColor(style.backgroundColor),
            type: ShadingType.CLEAR,
          },
  }
}

// ---------- Cell ----------

/**
 * Compile a column width value.
 *
 * If the value is a percentage string (e.g. `"30%"`), returns
 * a percentage width config. Otherwise returns `undefined`.
 *
 * @param width - — Column width value
 * @returns Width config or `undefined`
 *
 * @example
 * ```ts
 * compileColumnWidth('25%')
 * // => { size: 25, type: WidthType.PERCENTAGE }
 * ```
 */
export function compileColumnWidth(width: unknown) {
  if (typeof width === 'string' && width.endsWith('%')) {
    return {
      size: Number.parseFloat(width),
      type: WidthType.PERCENTAGE,
    }
  }
  return undefined
}

// ---------- Border ----------

/**
 * Compile a `DocxStyleRule` into `docx` paragraph-level options
 * (alignment, indent, spacing).
 *
 * @param style - — The resolved style rule
 * @returns Options object suitable for `new Paragraph(...)`
 *
 * @example
 * ```ts
 * compileParagraphStyle({ textAlign: 'center', marginTop: '10pt' })
 * // => { alignment: AlignmentType.CENTER, spacing: { before: 200, ... } }
 * ```
 */
export function compileParagraphStyle(style: DocxStyleRule) {
  const spacing = resolveSpacing(style)
  const indent = resolveIndent(style)

  return {
    alignment: compileAlignment(style.textAlign),
    indent,
    spacing,
  }
}

// ---------- Column width ----------

/**
 * Compile a `DocxStyleRule` into `docx` run-level options
 * (font, color, size, bold, italic, underline, etc.).
 *
 * Also merges any passthrough properties from `style.docx`.
 *
 * @param style - — The resolved style rule
 * @returns Options object suitable for `new TextRun(...)`
 *
 * @example
 * ```ts
 * compileTextStyle({ fontSize: 14, fontWeight: 'bold', color: '#ff0000' })
 * // => { size: 28, bold: true, color: 'ff0000' }
 * ```
 */
export function compileTextStyle(style: DocxStyleRule) {
  const result: Record<string, unknown> = {
    allCaps: style.allCaps,
    bold: style.fontWeight === 'bold' || Number(style.fontWeight) >= 600,
    color: normalizeColor(style.color),
    font: style.fontFamily,
    italics: style.fontStyle === 'italic',
    size: toPtHalf(style.fontSize),
    strike: style.strike,
    underline: style.underline ? {} : undefined,
  }

  // Merge docx escape hatch
  if (style.docx) {
    Object.assign(result, style.docx)
  }

  return result
}

// ---------- Helpers ----------

/**
 * Map the CSS-like text-align value to a `docx` `AlignmentType`.
 */
function compileAlignment(value: DocxStyleRule['textAlign']) {
  if (value === 'center') {
    return AlignmentType.CENTER
  }
  if (value === 'right') {
    return AlignmentType.RIGHT
  }
  if (value === 'justify') {
    return AlignmentType.JUSTIFIED
  }
  return AlignmentType.LEFT
}

/**
 * Map a CSS-like border-style string to a `docx` `BorderStyle` enum value.
 */
function compileBorderStyle(style?: import('../types/style').BorderStyle) {
  if (style === 'dashed') {
    return BorderStyle.DASHED
  }
  if (style === 'dotted') {
    return BorderStyle.DOTTED
  }
  if (style === 'double') {
    return BorderStyle.DOUBLE
  }
  if (style === 'none') {
    return BorderStyle.NONE
  }
  return BorderStyle.SINGLE
}

/** Convert line-height multiplier to twips. 1 → 240 (single-spacing). */
function compileLineHeight(value: DocxStyleRule['lineHeight']) {
  if (typeof value === 'number') {
    return Math.round(value * 240)
  }
  return toTwip(value)
}

/**
 * Compile a single border side into `docx` border options.
 */
function compileSingleBorder(rule: NonNullable<DocxStyleRule['border']>) {
  return {
    color: normalizeColor(rule.color),
    size: rule.width == null ? 1 : toTwip(rule.width),
    style: compileBorderStyle(rule.style),
  }
}

/**
 * Map the CSS-like vertical-align to a `docx` `VerticalAlign` enum value.
 */
function compileVerticalAlign(value: DocxStyleRule['verticalAlign']) {
  if (value === 'middle') {
    return VerticalAlign.CENTER
  }
  if (value === 'bottom') {
    return VerticalAlign.BOTTOM
  }
  return VerticalAlign.TOP
}

/** Strip leading `#` from a color string (docx expects raw hex). */
function normalizeColor(color?: string) {
  return color?.replace(/^#/, '')
}

/**
 * Resolve paragraph indent from a `DocxStyleRule`.
 *
 * Handles margin shorthand, explicit left/right margin, and text-indent.
 */
function resolveIndent(style: DocxStyleRule) {
  const shorthand =
    style.margin == null ? undefined : parseShorthandTwip(style.margin)

  const left = toTwip(style.marginLeft) ?? shorthand?.left
  const right = toTwip(style.marginRight) ?? shorthand?.right
  const firstLine = toTwip(style.textIndent)

  if (left == null && right == null && firstLine == null) {
    return undefined
  }
  return { firstLine, left, right }
}

/**
 * Resolve paragraph spacing from a `DocxStyleRule`.
 *
 * Handles margin shorthand for top/bottom spacing and line height.
 */
function resolveSpacing(style: DocxStyleRule) {
  const shorthand =
    style.margin == null ? undefined : parseShorthandTwip(style.margin)

  const before = toTwip(style.marginTop) ?? shorthand?.top
  const after = toTwip(style.marginBottom) ?? shorthand?.bottom
  const line = compileLineHeight(style.lineHeight)

  if (before == null && after == null && line == null) {
    return undefined
  }
  return { after, before, line }
}
