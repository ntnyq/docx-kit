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

// ---------- Unit conversion constants ----------

/**
 * Points per inch — standard typographic convention.
 *
 * All other conversions derive from this and the DPI assumption.
 */
const POINTS_PER_INCH = 72

/**
 * Pixels per inch at 96 DPI (standard screen resolution).
 */
const PIXELS_PER_INCH = 96

/**
 * Pixels per point at 96 DPI.
 *
 * 96 px/in ÷ 72 pt/in = 4/3 ≈ 1.333 px/pt.
 */
const PX_PER_POINT = PIXELS_PER_INCH / POINTS_PER_INCH // ≈ 1.333

/**
 * Points per pixel at 96 DPI.
 *
 * 72 pt/in ÷ 96 px/in = 0.75 pt/px.
 */
const POINTS_PER_PX = POINTS_PER_INCH / PIXELS_PER_INCH // = 0.75

/**
 * Points per millimeter.
 *
 * 72 pt/in ÷ 25.4 mm/in ≈ 2.835 pt/mm.
 */
const POINTS_PER_MM = POINTS_PER_INCH / 25.4 // ≈ 2.835

/**
 * Points per centimeter.
 *
 * POINTS_PER_MM × 10.
 */
const POINTS_PER_CM = POINTS_PER_MM * 10 // ≈ 28.346

/**
 * Twips per point — the fundamental Word OpenXML spacing unit.
 *
 * 1 twip = 1/20 of a point. All margins, indents, and spacing
 * values in Word are expressed in twips.
 */
const TWIPS_PER_POINT = 20

/**
 * Twips per pixel at 96 DPI.
 *
 * POINTS_PER_PX × TWIPS_PER_POINT = 0.75 × 20 = 15.
 */
const TWIPS_PER_PX = POINTS_PER_PX * TWIPS_PER_POINT // = 15

/**
 * Half-points per point (used for font sizes in Word OpenXML).
 *
 * Word stores font sizes in half-points: 12pt → 24 half-points.
 */
const HALF_POINTS_PER_POINT = 2

/**
 * Half-points per pixel at 96 DPI.
 *
 * POINTS_PER_PX × HALF_POINTS_PER_POINT = 0.75 × 2 = 1.5.
 */
const HALF_POINTS_PER_PX = POINTS_PER_PX * HALF_POINTS_PER_POINT // = 1.5

// ---------- Public API ----------

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
    return Math.round(value * HALF_POINTS_PER_POINT)
  }

  const n = Number.parseFloat(value)
  if (value.endsWith('pt')) {
    return Math.round(n * HALF_POINTS_PER_POINT)
  }
  if (value.endsWith('px')) {
    return Math.round(n * HALF_POINTS_PER_PX)
  }
  if (value.endsWith('mm')) {
    return Math.round(n * POINTS_PER_MM * HALF_POINTS_PER_POINT)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * POINTS_PER_CM * HALF_POINTS_PER_POINT)
  }
  if (value.endsWith('in')) {
    return Math.round(n * POINTS_PER_INCH * HALF_POINTS_PER_POINT)
  }

  return Math.round(n * HALF_POINTS_PER_POINT)
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
    return Math.round(n * PX_PER_POINT)
  }
  if (value.endsWith('mm')) {
    return Math.round(n * POINTS_PER_MM * PX_PER_POINT)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * POINTS_PER_CM * PX_PER_POINT)
  }
  if (value.endsWith('in')) {
    return Math.round(n * PIXELS_PER_INCH)
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
    return Math.round(value * TWIPS_PER_POINT)
  }

  const n = Number.parseFloat(value)
  if (value.endsWith('pt')) {
    return Math.round(n * TWIPS_PER_POINT)
  }
  if (value.endsWith('px')) {
    return Math.round(n * TWIPS_PER_PX)
  }
  if (value.endsWith('mm')) {
    return Math.round(n * POINTS_PER_MM * TWIPS_PER_POINT)
  }
  if (value.endsWith('cm')) {
    return Math.round(n * POINTS_PER_CM * TWIPS_PER_POINT)
  }
  if (value.endsWith('in')) {
    return Math.round(n * POINTS_PER_INCH * TWIPS_PER_POINT)
  }

  return Math.round(n * TWIPS_PER_POINT)
}
