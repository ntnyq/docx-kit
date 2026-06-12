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
 * CSS-like length value.
 *
 * Supports bare numbers, and explicit unit strings:
 * `%`, `cm`, `in`, `mm`, `pt`, `px`.
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
