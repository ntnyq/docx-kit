import { DocxKitError } from '@docxkit/errors'
import type { StyleSheet } from '@docxkit/types'

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
