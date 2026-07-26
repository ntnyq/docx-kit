/**
 * Browser URL plugin loader.
 *
 * Loads a plugin from a remote URL via dynamic `import()`.
 * The plugin code MUST be bundled as an ES module
 * and export a valid DocxPlugin (either as the default export
 * or as a named export matching the plugin name).
 *
 * This file is browser only. Import it from `docx-kit/loader/browser`.
 *
 * SECURITY WARNING: Loading plugins from arbitrary URLs
 * is inherently risky. Always validate plugins from trusted sources
 * and use the {@link PluginSecurityPolicy} to restrict origins.
 *
 * @module loader/sources/url
 */

import { DocxKitError } from '@docxkit/core'
import { isDocxPlugin } from '../utils'
import { resolveManifest } from './options'

import type { DocxPlugin } from '@docxkit/core'
import type { ExternalPluginLoadOptions } from './options'

/**
 * Load a plugin from a remote URL (browser).
 *
 * @param url - — The URL to load the plugin ES module from
 * @param options - — Optional loading options (manifest override)
 * @param options.manifest - — Pre-resolved manifest (avoids re-reading)
 * @returns The loaded plugin and optionally a manifest
 * @throws {DocxKitError} `PLUGIN_LOAD_FAILED` if import fails
 */
export async function loadUrlPlugin(
  url: string,
  options?: ExternalPluginLoadOptions,
): Promise<{
  manifest: Awaited<ReturnType<typeof resolveManifest>>
  plugin: DocxPlugin
}> {
  const manifest = await resolveManifest(options)

  let mod: DocxPlugin | { default?: unknown }
  try {
    mod = (await import(/* @vite-ignore */ url)) as
      DocxPlugin | { default?: unknown }
  } catch (error) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Failed to import plugin from URL: "${url}"`,
      error,
    )
  }

  // Extract the plugin — check for default export first, then named
  const exported = mod as { default?: unknown }
  const plugin = exported.default ?? mod

  if (!isDocxPlugin(plugin)) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Module loaded from "${url}" does not export a valid DocxPlugin`,
    )
  }

  // If a manifest was provided, verify plugin name matches
  if (manifest && plugin.name !== manifest.plugin.name) {
    console.warn(
      `[docx-kit] Plugin name mismatch: manifest declares "${manifest.plugin.name}" `
        + `but loaded plugin has "${plugin.name}"`,
    )
  }

  return {
    manifest,
    plugin: plugin as DocxPlugin,
  }
}
