/**
 * Local file plugin loader.
 *
 * Loads a plugin from a local file path or directory.
 * Platform-specific implementations handle the actual loading
 * (filesystem in Node.js, fetch in browser).
 *
 * This module provides the shared interface and a cross-platform
 * dispatch mechanism.
 *
 * @module loader/sources/local
 */

import { DocxKitError } from '../../errors'
import type { DocxPlugin } from '../../types/plugin'
import type { PluginManifest } from '../manifest'

/**
 * Load a plugin from a local file path.
 *
 * In Node.js, dynamically imports from the filesystem.
 * In the browser, uses `fetch()` for `file://` URLs or
 * dynamic import for same-origin paths.
 *
 * @param path - — The file path to load
 * @param options - — Optional loading options
 * @param options.manifest - — Pre-resolved manifest (avoids re-reading)
 * @returns The loaded plugin and its manifest
 * @throws {DocxKitError} `PLUGIN_LOAD_FAILED` if import fails
 */
export async function loadLocalPlugin(
  path: string,
  options?: { manifest?: PluginManifest },
): Promise<{ manifest: PluginManifest; plugin: DocxPlugin }> {
  let mod: DocxPlugin | { default?: unknown }
  try {
    // Dynamic import works for both Node.js (file:// URLs)
    // and browser (same-origin paths). The caller is responsible
    // for providing a valid absolute path or URL.
    mod = (await import(path)) as DocxPlugin | { default?: unknown }
  } catch (err) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Failed to import plugin from local path: "${path}"`,
      err,
    )
  }

  // Extract the plugin
  const exported = mod as { default?: unknown }
  const plugin = exported.default ?? mod

  if (!isDocxPlugin(plugin)) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Module loaded from "${path}" does not export a valid DocxPlugin`,
    )
  }

  return {
    manifest: options?.manifest ?? (null as unknown as PluginManifest),
    plugin: plugin as DocxPlugin,
  }
}

/**
 * Type guard: check if a value is a valid DocxPlugin.
 */
function isDocxPlugin(value: unknown): value is DocxPlugin {
  return (
    !!value
    && typeof value === 'object'
    && typeof (value as DocxPlugin).name === 'string'
    && typeof (value as DocxPlugin).render === 'function'
  )
}
