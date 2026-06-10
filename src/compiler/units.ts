/**
 * Word OpenXML unit conversion utilities.
 *
 * All bare numbers are interpreted according to the property they represent:
 * - `fontSize` (toPtHalf) → **pt**
 * - spacing / indent / margin (toTwip) → **pt**
 * - image size (toPx) → **px**
 *
 * @module compiler/units
 */

/**
 * Parse a CSS margin/padding shorthand string into top/right/bottom/left
 * values expressed in **twips** (1/20 pt).
 *
 * Supports 1-value, 2-value, and 4-value shorthand.
 *
 * @param value - — Single number, unit string, or CSS shorthand
 * @returns `{ top, right, bottom, left }` values in twips, or `undefined`
 *
 * @example
 * ```ts
 * parseShorthandTwip('10pt 20pt')
 * // => { top: 200, right: 400, bottom: 200, left: 400 }
 *
 * parseShorthandTwip('5pt 10pt 15pt 20pt')
 * // => { top: 100, right: 200, bottom: 300, left: 400 }
 * ```
 */
export function parseShorthandTwip(
  value: number | string | undefined,
): { bottom: number; left: number; right: number; top: number } | undefined {
  if (value == null) {
    return undefined
  }
  const parts = String(value).trim().split(/\s+/)
  if (parts.length === 1) {
    const v = toTwip(parts[0]) ?? 0
    return { bottom: v, left: v, right: v, top: v }
  }
  if (parts.length === 2) {
    const [tb, lr] = parts.map(p => toTwip(p) ?? 0)
    return { bottom: tb, left: lr, right: lr, top: tb }
  }
  if (parts.length === 4) {
    const [top, right, bottom, left] = parts.map(p => toTwip(p) ?? 0)
    return { bottom, left, right, top }
  }
  return undefined
}

/**
 * Convert a font-size value to Word **half-points** (used by the `docx`
 * library's `size` field on text runs).
 *
 * Bare number is treated as **pt** (e.g. `12` → 24 half-points).
 *
 * @param value - — Number (pt) or unit string (`"12pt"`, `"16px"`, `"4mm"`, etc.)
 * @returns Half-point integer, or `undefined` if input is null/undefined
 *
 * @example
 * ```ts
 * toPtHalf(12)        // 24   (12pt × 2)
 * toPtHalf('14pt')    // 28
 * toPtHalf('16px')    // 24   (16px ≈ 12pt)
 * ```
 */
export function toPtHalf(
  value: number | string | undefined,
): number | undefined {
  if (value == null) {
    return undefined
  }
  if (typeof value === 'number') {
    return Math.round(value * 2)
  }

  const n = Number.parseFloat(value)
  if (value.endsWith('pt')) {
    return Math.round(n * 2)
  }
  if (value.endsWith('px')) {
    return Math.round(n * 1.5)
  }
  if (value.endsWith('mm')) {
    return Math.round(n * 5.66929)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * 56.6929)
  }
  if (value.endsWith('in')) {
    return Math.round(n * 144)
  }

  return Math.round(n * 2)
}

/**
 * Convert a length value to **pixels** (used for image dimensions).
 *
 * Bare number is treated as **px**.
 *
 * @param value - — Number (px) or unit string (`"100px"`, `"72pt"`, `"5cm"`, etc.)
 * @returns Pixel integer, or `undefined` if input is null/undefined
 *
 * @example
 * ```ts
 * toPx(200)          // 200
 * toPx('150px')      // 150
 * toPx('72pt')       // 96   (72pt × 1.333)
 * ```
 */
export function toPx(value: number | string | undefined): number | undefined {
  if (value == null) {
    return undefined
  }
  if (typeof value === 'number') {
    return value
  }

  const n = Number.parseFloat(value)
  if (value.endsWith('px')) {
    return n
  }
  if (value.endsWith('pt')) {
    return Math.round(n * 1.333333)
  }
  if (value.endsWith('mm')) {
    return Math.round(n * 3.779527)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * 37.79527)
  }
  if (value.endsWith('in')) {
    return Math.round(n * 96)
  }

  return n
}

/**
 * Convert a length value to **twips** (1/20 of a point).
 *
 * Used for spacing, indent, and margin values in Word OpenXML.
 * Bare number is treated as **pt**.
 *
 * @param value - — Number (pt) or unit string (`"10pt"`, `"12px"`, `"5mm"`, etc.)
 * @returns Twip integer, or `undefined` if input is null/undefined
 *
 * @example
 * ```ts
 * toTwip(12)         // 240  (12pt × 20)
 * toTwip('1in')      // 1440 (1 inch × 1440)
 * toTwip('10mm')     // 567  (10mm × 56.7)
 * ```
 */
export function toTwip(value: number | string | undefined): number | undefined {
  if (value == null) {
    return undefined
  }
  if (typeof value === 'number') {
    return Math.round(value * 20)
  }

  const n = Number.parseFloat(value)
  if (value.endsWith('pt')) {
    return Math.round(n * 20)
  }
  if (value.endsWith('px')) {
    return Math.round(n * 15)
  } // 96 dpi: 1px ≈ 0.75pt
  if (value.endsWith('mm')) {
    return Math.round(n * 56.6929)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * 566.929)
  }
  if (value.endsWith('in')) {
    return Math.round(n * 1440)
  }

  return Math.round(n * 20)
}
