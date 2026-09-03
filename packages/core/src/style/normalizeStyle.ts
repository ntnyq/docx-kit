/**
 * Style normalization — merge multiple style sources into one resolved rule.
 *
 * Implements the CSS-like cascade:
 * `base → className(s) [with extends resolution] → inline style → theme tokens`
 *
 * @module style/normalizeStyle
 */

import { DocxKitError } from '@docxkit/types'
import { resolveThemeTokens } from './theme'
import type {
  DocxStyleRule,
  DocxTheme,
  StyleSheet,
  StyleSheetEntry,
} from '@docxkit/types'

/**
 * Options for resolving a style cascade.
 */
export interface ResolveStyleOptions {
  /**
   * Base style (typically element defaults).
   */
  base?: DocxStyleRule
  /**
   * Class name(s) to resolve from the stylesheet.
   */
  className?: string | string[]
  /**
   * Inline style overrides (highest priority).
   */
  inline?: DocxStyleRule
  /**
   * The stylesheet to resolve class names against.
   */
  styles?: StyleSheet
  /**
   * Theme for token resolution (color, font, spacing tokens).
   */
  theme?: DocxTheme
}

/**
 * Merge styles in priority order (lowest → highest):
 * `base → className(s) [with extends] → inline → theme tokens`.
 *
 * - **extends**: Each className is resolved recursively — if a stylesheet
 *   entry has an `extends` property, its parent class(es) are resolved first,
 *   then the entry's own properties override them. Circular references are
 *   detected and rejected.
 * - **theme tokens**: After the cascade merge, any string value starting with
 *   `$` (e.g. `$colors.primary`, `$fonts.body`) is resolved against the theme.
 *
 * @returns A new merged `DocxStyleRule` (does not mutate inputs).
 *
 * @example
 * ```ts
 * const resolved = resolveStyle({
 *   base:      { fontSize: 12 },
 *   className: ['body', 'blue'],
 *   inline:    { fontWeight: 'bold' },
 *   styles:    { body: { lineHeight: 1.5 }, blue: { color: '$colors.primary' } },
 *   theme:     { colors: { primary: '#1a56db' } },
 * })
 * // => { fontSize: 12, lineHeight: 1.5, color: '#1a56db', fontWeight: 'bold' }
 * ```
 */
export function resolveStyle(options: ResolveStyleOptions): DocxStyleRule {
  const classNames = Array.isArray(options.className)
    ? options.className
    : options.className
      ? options.className.split(/\s+/).filter(Boolean)
      : []

  const classStyles = classNames.map(name =>
    resolveStyleClass(name, options.styles, new Set()),
  )

  // Merge: base → class(es) → inline
  const merged = Object.assign({}, options.base, ...classStyles, options.inline)

  // Resolve theme tokens as the final pass
  return resolveThemeTokens(merged, options.theme)
}

/**
 * Resolve a single stylesheet class name with `extends` inheritance.
 *
 * Walks the `extends` chain left-to-right (parent styles first, child overrides on top),
 * detecting circular references. The `extends` metadata is stripped from the result.
 */
function resolveStyleClass(
  name: string,
  styles: StyleSheet | undefined,
  visited: Set<string>,
): DocxStyleRule {
  if (visited.has(name)) {
    throw new DocxKitError(
      'STYLE_CIRCULAR_EXTENDS',
      `Circular extends detected for style class "${name}". `
        + `Chain: ${[...visited, name].join(' → ')}`,
    )
  }

  const entry = styles?.[name]
  if (!entry) {
    // Only throw if styles are explicitly provided — allows optional stylesheets
    if (styles) {
      throw new DocxKitError(
        'STYLE_UNKNOWN_CLASS',
        `Style class not found: "${name}"`,
      )
    }
    return {}
  }

  visited.add(name)

  // Extract extends metadata and the actual style rule
  const extendsValue = (entry as StyleSheetEntry).extends
  const { extends: _, ...styleRule } = entry as StyleSheetEntry

  let resolved: DocxStyleRule = styleRule

  if (extendsValue) {
    const parents = Array.isArray(extendsValue) ? extendsValue : [extendsValue]
    const parentStyles = parents.map(parent =>
      resolveStyleClass(parent, styles, new Set(visited)),
    )
    // Merge: parent styles as base, own properties override
    resolved = Object.assign({}, ...parentStyles, styleRule)
  }

  visited.delete(name)
  return resolved
}
