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
  HighlightColor,
  ShadingType,
  VerticalAlign,
  WidthType,
} from 'docx'
import { parseShorthandTwip, toPtHalf, toTwip } from './units'
import type {
  CellStyleRule,
  DocxStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/types'

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
export function compileBorder(style: CellStyleRule | ParagraphStyleRule) {
  const b = style.border
  const bt = style.borderTop ?? b
  const br = style.borderRight ?? b
  const bb = style.borderBottom ?? b
  const bl = style.borderLeft ?? b

  if (!bt && !br && !bb && !bl) {
    return undefined
  }

  return {
    bottom: bb ? compileBorderRule(bb!) : undefined,
    left: bl ? compileBorderRule(bl!) : undefined,
    right: br ? compileBorderRule(br!) : undefined,
    top: bt ? compileBorderRule(bt!) : undefined,
  }
}

// ---------- Border ----------

/**
 * Compile a single border side into `docx` border options.
 */
export function compileBorderRule(rule: NonNullable<DocxStyleRule['border']>) {
  return {
    color: normalizeColor(rule.color),
    size: rule.width == null ? 1 : toTwip(rule.width),
    style: compileBorderStyle(rule.style),
  }
}

// ---------- Cell ----------

/**
 * Compile a `DocxStyleRule` into `docx` table cell options
 * (vertical alignment, margins, shading).
 *
 * Applies sensible default cell padding (5 pt left/right, 2 pt top/bottom)
 * when no margins are explicitly set.
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
export function compileCellStyle(style: CellStyleRule) {
  const shorthand =
    style.margin == null ? undefined : parseShorthandTwip(style.margin)

  // Default cell padding: 100 twips (5 pt) left/right, 0 top/bottom.
  // These are overridden when any margin value is explicitly provided.
  const bottom = toTwip(style.marginBottom) ?? shorthand?.bottom
  const left = toTwip(style.marginLeft) ?? shorthand?.left
  const right = toTwip(style.marginRight) ?? shorthand?.right
  const top = toTwip(style.marginTop) ?? shorthand?.top

  const hasAnyMargin =
    bottom != null || left != null || right != null || top != null

  return {
    verticalAlign: compileVerticalAlign(style.verticalAlign),
    margins: hasAnyMargin
      ? {
          bottom: bottom ?? 0,
          left: left ?? 0,
          right: right ?? 0,
          top: top ?? 0,
        }
      : {
          bottom: 0,
          left: 100,
          right: 100,
          top: 0,
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

// ---------- Border ----------

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

/** Compile paragraph-level borders (separate from cell borders). */
export function compileParagraphBorder(style: ParagraphStyleRule) {
  const borders = compileBorder(style)
  if (!borders) {
    return {}
  }
  return { border: borders }
}

// ---------- Column width ----------

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
export function compileParagraphStyle(style: ParagraphStyleRule) {
  const spacing = resolveSpacing(style)
  const indent = resolveIndent(style)

  return {
    alignment: compileAlignment(style.textAlign),
    ...compileParagraphBorder(style),
    indent,
    keepLines: style.keepLines,
    keepNext: style.keepNext,
    pageBreakBefore: style.pageBreakBefore,
    spacing,
  }
}

// ---------- Helpers ----------

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
export function compileTextStyle(style: TextStyleRule) {
  const result: Record<string, unknown> = {
    allCaps: style.allCaps,
    bold: style.fontWeight === 'bold' || Number(style.fontWeight) >= 600,
    characterSpacing: style.characterSpacing,
    color: normalizeColor(style.color),
    font: style.fontFamily,
    highlight: compileHighlight(style.highlight),
    italics: style.fontStyle === 'italic',
    size: toPtHalf(style.fontSize),
    smallCaps: style.smallCaps,
    strike: style.strike,
    subScript: style.subScript,
    superScript: style.superScript,
    underline: style.underline ? {} : undefined,
  }

  // Merge docx escape hatch
  if (style.docx) {
    Object.assign(result, style.docx)
  }

  return result
}

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
function compileBorderStyle(style?: import('@docxkit/types').BorderStyle) {
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

/**
 * Map a highlight color name to `docx` `HighlightColor`.
 */
function compileHighlight(color?: import('@docxkit/types').HighlightColor) {
  if (!color || color === 'none') {
    return undefined
  }
  return (HighlightColor as Record<string, unknown>)[color.toUpperCase()]
}

/** Convert line-height multiplier to twips. 1 → 240 (single-spacing). */
function compileLineHeight(value: DocxStyleRule['lineHeight']) {
  if (typeof value === 'number') {
    return Math.round(value * 240)
  }
  return toTwip(value)
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
