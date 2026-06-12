/**
 * Shared loader utilities.
 *
 * Utility functions used across multiple source loaders
 * (npm, url, local, inline).
 *
 * @module loader/utils
 */

import type { DocxPlugin } from '@docxkit/core'

/**
 * Type guard: check if a value is a valid DocxPlugin.
 *
 * A valid plugin must be an object with a `name` string property
 * and a `render` function.
 *
 * @param value - — Value to check
 * @returns `true` if the value has a `name` string and `render` function
 */
export function isDocxPlugin(value: unknown): value is DocxPlugin {
  return (
    !!value
    && typeof value === 'object'
    && typeof (value as DocxPlugin).name === 'string'
    && typeof (value as DocxPlugin).render === 'function'
  )
}
