/**
 * Built-in theme tokens for docx-kit.
 *
 * Themes provide semantic design tokens (colors, fonts, spacing) that can be
 * referenced in styles via `$category.key` syntax. Use them as the `theme`
 * property in `DocxKitConfig`.
 *
 * @module themes
 */

import { minimalTheme } from './minimal'
import { oceanTheme } from './ocean'
import { warmTheme } from './warm'
import type { DocxTheme } from '../types/document'

export type { DocxTheme } from '../types/document'

/** All built-in themes, keyed by ID. */
export const BUILTIN_THEMES: ReadonlyMap<string, DocxTheme> = new Map([
  [minimalTheme.id, minimalTheme],
  [oceanTheme.id, oceanTheme],
  [warmTheme.id, warmTheme],
])

/** Ordered list of built-in themes (for UI selectors). */
export const THEME_LIST: readonly DocxTheme[] = [
  minimalTheme,
  oceanTheme,
  warmTheme,
]

/**
 * Look up a built-in theme by ID.
 *
 * @param id - — Theme identifier (`"minimal"`, `"ocean"`, or `"warm"`)
 * @returns The matching theme, or `undefined` if not found
 *
 * @example
 * ```ts
 * import { createDocx, useTheme } from 'docx-kit'
 * const doc = createDocx({ theme: useTheme('ocean') })
 * ```
 */
export function useTheme(id: string): DocxTheme | undefined {
  return BUILTIN_THEMES.get(id)
}

export { minimalTheme, oceanTheme, warmTheme }
