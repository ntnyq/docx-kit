/**
 * Theme token resolution — process `$category.key` references in style values.
 *
 * @module style/theme
 */

import type { DocxTheme } from '@docxkit/types'

/**
 * Resolve theme token references in a style rule or selected layout fields.
 * Callers must exclude document text and other literal content.
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
 * @param value - Style or layout data to resolve tokens in
 * @param theme - — The theme to resolve against
 * @returns Data with tokens replaced by their values, without mutating the input
 */
export function resolveThemeTokens<T>(value: T, theme?: DocxTheme): T {
  if (!theme) {
    return value
  }

  return resolveTokenValue(value, theme) as T
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object') {
    return false
  }

  const prototype = Object.getPrototypeOf(value)
  return prototype === null || prototype === Object.prototype
}

/**
 * Parse `$category.key` and look up the value in the theme.
 *
 * @param token - Theme token in `$category.key` form
 * @param theme - Theme containing the token category maps
 * @returns The token value, or `undefined` if the token or key is not recognized
 */
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

/**
 * Recursively resolve tokens in nested style values such as borders and tab
 * stops. Style rules are plain data, but the prototype check keeps opaque
 * class instances in the `docx` escape hatch untouched.
 */
function resolveTokenValue(value: unknown, theme: DocxTheme): unknown {
  if (typeof value === 'string' && value.startsWith('$')) {
    return resolveSingleToken(value, theme) ?? value
  }

  if (Array.isArray(value)) {
    return value.map(item => resolveTokenValue(item, theme))
  }

  if (isPlainObject(value)) {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [
        key,
        resolveTokenValue(item, theme),
      ]),
    )
  }

  return value
}
