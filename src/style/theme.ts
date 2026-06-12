/**
 * Theme token resolution — process `$category.key` references in style values.
 *
 * @module style/theme
 */

import type { DocxTheme } from '../types/document'
import type { DocxStyleRule } from '../types/style'

/**
 * Resolve all theme token references in a style rule.
 *
 * Token syntax: `$category.key`
 * - `$colors.primary` → `theme.colors?.['primary']`
 * - `$fonts.heading`   → `theme.fonts?.['heading']`
 * - `$spacing.md`      → `theme.spacing?.['md']`
 * - `$fontSize.lg`      → `theme.fontSize?.['lg']`
 *
 * Non-token values pass through unchanged. Unknown token references
 * are left as-is (no error — styled text may use $ as a literal).
 *
 * @param rule - — The style rule to resolve tokens in
 * @param theme - — The theme to resolve against
 * @returns A new `DocxStyleRule` with tokens replaced by their values
 */
export function resolveThemeTokens(
  rule: DocxStyleRule,
  theme?: DocxTheme,
): DocxStyleRule {
  if (!theme) {
    return rule
  }

  const resolved: Record<string, unknown> = { ...rule }

  for (const key of Object.keys(rule)) {
    const value = (rule as Record<string, unknown>)[key]
    if (typeof value === 'string' && value.startsWith('$')) {
      const resolvedValue = resolveSingleToken(value, theme)
      if (resolvedValue !== undefined) {
        resolved[key] = resolvedValue
      }
    }
  }

  return resolved as DocxStyleRule
}

/** Parse `$category.key` and look up the value in the theme. */
function resolveSingleToken(token: string, theme: DocxTheme): unknown {
  const re = /^\$(colors|fonts|fontSize|spacing)\.(.+)$/
  const match = token.match(re)
  if (!match) {
    return undefined
  }

  const [, category, key] = match
  const map = theme[category as keyof DocxTheme]
  if (map && typeof map === 'object') {
    return (map as Record<string, unknown>)[key]
  }
  return undefined
}
