/**
 * Style normalization — merge multiple style sources into one resolved rule.
 *
 * Implements the CSS-like cascade:
 * `defaults → className(s) → inline style`
 *
 * @module style/normalizeStyle
 */

import type { DocxStyleRule, StyleSheet } from '../types/style'

/**
 * Merge styles in priority order (lowest → highest):
 * `base` → `className`(s) → `inline style`.
 *
 * - `base` — Typically a default style for the element type
 * - `className` — Single string, space-separated string, or array of class names
 * - `inline` — Inline style override on the node itself
 * - `styles` — The stylesheet map to resolve class names against
 *
 * @returns A new merged `DocxStyleRule` (does not mutate inputs).
 *
 * @example
 * ```ts
 * const resolved = resolveStyle({
 *   base:      { fontSize: 12 },
 *   className: ['body', 'blue'],
 *   inline:    { fontWeight: 'bold' },
 *   styles:    { body: { lineHeight: 1.5 }, blue: { color: '#00f' } },
 * })
 * // => { fontSize: 12, lineHeight: 1.5, color: '#00f', fontWeight: 'bold' }
 * ```
 */
export function resolveStyle(options: {
  base?: DocxStyleRule
  className?: string | string[]
  inline?: DocxStyleRule
  styles?: StyleSheet
}): DocxStyleRule {
  const classNames = Array.isArray(options.className)
    ? options.className
    : options.className
      ? options.className.split(/\s+/)
      : []

  const classStyles = classNames
    .map(name => options.styles?.[name])
    .filter((s): s is DocxStyleRule => s != null)

  return Object.assign({}, options.base, ...classStyles, options.inline)
}
