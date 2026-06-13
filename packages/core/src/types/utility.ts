/**
 * Generic utility types used throughout docx-kit.
 *
 * @module types/utility
 */

/** Generic dictionary with string keys. */
export type Dict<T = unknown> = Record<string, T>

/** Hexadecimal CSS color string, e.g. `"#ff0000"`. */
export type HexColor = `#${string}`

/**
 * A literal union type that still allows arbitrary string values.
 * Useful for autocomplete-friendly APIs with extensible values.
 *
 * @template T — The literal subtype (e.g. `'Arial' | 'Calibri'`)
 * @template U — The base type, defaults to `string`
 */
export type LiteralUnion<T extends U, U = string> = T | (U & {})

/**
 * A value that might be synchronous or wrapped in a Promise.
 *
 * @template T — The inner value type
 */
export type MaybePromise<T> = Promise<T> | T

/**
 * Augments a base value type with theme token references.
 *
 * Lets users write `$colors.primary` in any style field that accepts the
 * base type — the token is resolved at compile time against the document
 * theme.
 *
 * @example
 * ```ts
 * const rule: DocxStyleRule = {
 *   // Plain UnitValue:
 *   fontSize: 12,
 *   // Theme token resolved against `theme.fontSize.lg`:
 *   marginBottom: '$spacing.lg',
 *   // `StyleToken<UnitValue>` would also accept both:
 *   marginTop: '$spacing.md' as StyleToken<UnitValue>,
 * }
 * ```
 */
export type StyleToken<T extends number | string> = T | ThemeToken

/**
 * Supported theme token categories (must match {@link resolveSingleToken}).
 */
export type StyleTokenCategory = '$colors' | '$fonts' | '$fontSize' | '$spacing'

/**
 * A `$category.key` reference to a theme token, e.g. `"$colors.primary"`,
 * `"$fontSize.lg"`, `"$spacing.xl"`. Resolved at compile time by
 * {@link resolveThemeTokens}.
 *
 * Use {@link StyleToken} to combine this with a concrete value type.
 */
export type ThemeToken = `${StyleTokenCategory}.${string}`

/**
 * CSS-like length value.
 *
 * Supports bare numbers, explicit unit strings
 * (`%`, `cm`, `in`, `mm`, `pt`, `px`), and theme token references
 * (e.g. `"$spacing.lg"`, `"$fontSize.base"`).
 *
 * Bare-number conventions vary by context:
 * - `fontSize` → interpreted as `pt`
 * - spacing / indent → interpreted as `pt`
 * - image size → interpreted as `px`
 */
export type UnitValue =
  | `${number}%`
  | `${number}cm`
  | `${number}in`
  | `${number}mm`
  | `${number}pt`
  | `${number}px`
  | number
  | ThemeToken
