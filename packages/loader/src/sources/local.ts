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

import { DocxKitError } from '@docxkit/core'
import { resolvePluginExport } from '../utils'
import { resolveManifest } from './options'

import type { DocxPlugin } from '@docxkit/core'
import type { ExternalPluginLoadOptions } from './options'

/**
 * Load a plugin from a local file path.
 *
 * In Node.js, dynamically imports from the filesystem.
 * In the browser, uses `fetch()` for `file://` URLs or
 * dynamic import for same-origin paths.
 *
 * @param path - — The file path to load
 * @param options - — Optional loading options
 * @param options.manifest - — Manifest resolved before importing code
 * @param options.authorizeManifest - — Pre-import manifest authorization
 * @returns The loaded plugin and its manifest
 * @throws {DocxKitError} `PLUGIN_LOAD_FAILED` if import fails
 */
export async function loadLocalPlugin(
  path: string,
  options?: ExternalPluginLoadOptions,
): Promise<{
  manifest: Awaited<ReturnType<typeof resolveManifest>>
  plugin: DocxPlugin
}> {
  const manifest = await resolveManifest(options)

  let mod: DocxPlugin | { default?: unknown }
  try {
    // Dynamic import works for both Node.js (file:// URLs)
    // and browser (same-origin paths). The caller is responsible
    // for providing a valid absolute path or URL. Bundlers must preserve
    // this runtime-resolved specifier instead of trying to analyze it.
    mod = (await import(/* @vite-ignore */ path)) as
      | DocxPlugin
      | {
          default?: unknown
        }
  } catch (error) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Failed to import plugin from local path: "${path}"`,
      error,
    )
  }

  const plugin = await resolvePluginExport(mod, path, manifest?.plugin.name)

  return {
    manifest,
    plugin,
  }
}
