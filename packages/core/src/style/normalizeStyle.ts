/**
 * Style normalization — merge multiple style sources into one resolved rule.
 *
 * Implements the CSS-like cascade:
 * `base → className(s) [with extends resolution] → inline style → theme tokens`
 *
 * @module style/normalizeStyle
 */

import { createDocxKitError } from '@docxkit/errors'
import { resolveThemeTokens } from './theme'
import type {
  DocxStyleRule,
  DocxTheme,
  StyleSheet,
  StyleSheetEntry,
} from '@docxkit/types'

const STYLE_ALIASES = [
  ['margin', ['marginBottom', 'marginLeft', 'marginRight', 'marginTop']],
  ['bold', ['fontWeight']],
  ['fontWeight', ['bold']],
  ['italic', ['fontStyle']],
  ['fontStyle', ['italic']],
  ['characterSpacing', ['letterSpacing']],
  ['letterSpacing', ['characterSpacing']],
] as const

/**
 * Options for resolving a style cascade.
 */
export interface ResolveStyleOptions {
  /**
   * Base style or ordered default/inherited layers (lowest priority first).
   */
  base?: (DocxStyleRule | undefined)[] | DocxStyleRule
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

  const cache = new Map<string, DocxStyleRule>()
  const classStyles = classNames.map(name =>
    resolveStyleClass(name, options.styles, new Set(), cache),
  )

  // Merge: base → class(es) → inline
  const baseLayers = Array.isArray(options.base) ? options.base : [options.base]
  const merged = mergeStyleRules(...baseLayers, ...classStyles, options.inline)

  // Resolve theme tokens as the final pass
  return resolveThemeTokens(merged, options.theme)
}

/**
 * A higher layer's shorthand or alias replaces earlier declarations of the
 * same meaning. Merge each layer afterward to preserve same-layer precedence.
 * JSON keys must never invoke the legacy prototype setter.
 */
function mergeStyleRules(
  ...rules: (DocxStyleRule | undefined)[]
): DocxStyleRule {
  let merged: DocxStyleRule = {}
  for (const rule of rules) {
    if (!rule) {
      continue
    }
    for (const [property, aliases] of STYLE_ALIASES) {
      if (rule[property] !== undefined) {
        for (const alias of aliases) {
          delete merged[alias]
        }
      }
    }
    merged = Object.hasOwn(rule, '__proto__')
      ? Object.fromEntries([...Object.entries(merged), ...Object.entries(rule)])
      : Object.assign(merged, rule)
  }
  return merged
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
  cache: Map<string, DocxStyleRule>,
): DocxStyleRule {
  if (visited.has(name)) {
    throw createDocxKitError('STYLE_CIRCULAR_EXTENDS', {
      chain: [...visited, name],
      className: name,
    })
  }

  const cached = cache.get(name)
  if (cached) {
    return cached
  }

  const entry = styles && Object.hasOwn(styles, name) ? styles[name] : undefined
  if (!entry) {
    // Only throw if styles are explicitly provided — allows optional stylesheets
    if (styles) {
      throw createDocxKitError('STYLE_UNKNOWN_CLASS', { className: name })
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
      resolveStyleClass(parent, styles, visited, cache),
    )
    // Merge: parent styles as base, own properties override
    resolved = mergeStyleRules(...parentStyles, styleRule)
  }

  visited.delete(name)
  cache.set(name, resolved)
  return resolved
}
