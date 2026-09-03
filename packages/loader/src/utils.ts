/**
 * Shared loader utilities.
 *
 * Utility functions used across multiple source loaders
 * (npm, url, local, inline).
 *
 * @module loader/utils
 */

import { DocxKitError } from '@docxkit/core'
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

/**
 * Resolve the shared module contract: a default instance/factory, a matching
 * named export, or a single unambiguous named instance/factory (legacy scaffold).
 */
export async function resolvePluginExport(
  module: unknown,
  source: string,
  name?: string,
): Promise<DocxPlugin> {
  let candidate = module
  if (!isDocxPlugin(module) && typeof module === 'object' && module !== null) {
    const exports = module as Record<string, unknown>
    const named = Object.values(exports).filter(
      value => isDocxPlugin(value) || typeof value === 'function',
    )
    candidate =
      exports.default
      ?? (name && Object.hasOwn(exports, name) ? exports[name] : undefined)
      ?? (named.length === 1 ? named[0] : undefined)
  }
  try {
    const plugin: unknown =
      typeof candidate === 'function' ? await candidate() : candidate
    if (isDocxPlugin(plugin)) {
      return plugin
    }
  } catch (error) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Plugin factory from "${source}" failed`,
      error,
    )
  }
  throw new DocxKitError(
    'PLUGIN_LOAD_FAILED',
    `Module loaded from "${source}" does not export a valid DocxPlugin; use a default plugin instance or factory`,
  )
}
