/**
 * Node.js npm package plugin loader.
 *
 * Dynamically imports plugins from `node_modules` and
 * resolves their `docx-kit.plugin.json` manifests.
 *
 * This file is Node.js only. Import it from `docx-kit/loader/node`.
 *
 * @module loader/sources/npm
 */

import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { DocxKitError, validateManifest } from '@docxkit/core'
import { isDocxPlugin } from '../utils'

import type { DocxPlugin, PluginManifest } from '@docxkit/core'

/**
 * Load a plugin from an npm package in `node_modules`.
 *
 * Resolves the package path via `import.meta.resolve()`,
 * reads the `docx-kit.plugin.json` manifest, validates it,
 * then dynamically imports the plugin module.
 *
 * @param packageName - — npm package name (e.g. `"docx-kit-chart"`)
 * @param _options - — Optional loading options (reserved for future use)
 * @param _options.cwd - — Working directory override (reserved)
 * @returns The loaded plugin and its manifest
 * @throws {DocxKitError} `MANIFEST_MISSING` if manifest is not found
 * @throws {DocxKitError} `MANIFEST_INVALID` if manifest validation fails
 * @throws {DocxKitError} `PLUGIN_LOAD_FAILED` if import fails
 */
export async function loadNpmPlugin(
  packageName: string,
  _options?: { cwd?: string },
): Promise<{ manifest: PluginManifest; plugin: DocxPlugin }> {
  // Resolve the package root directory
  let pkgPath: string

  try {
    // Use import.meta.resolve for Node 20.6+
    const resolved = import.meta.resolve(packageName)
    pkgPath = fileURLToPath(resolved)

    // Walk up from the resolved path to find package root
    // import.meta.resolve typically points to the main entry
    // e.g. .../node_modules/docx-kit-chart/dist/index.js
    // We need to walk up to the package root
    pkgPath = resolvePackageRoot(pkgPath)
  } catch {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `npm package "${packageName}" could not be resolved. Is it installed?`,
    )
  }

  // Read and validate the manifest
  const manifestPath = `${pkgPath}/docx-kit.plugin.json`

  if (!existsSync(manifestPath)) {
    throw new DocxKitError(
      'MANIFEST_MISSING',
      `Plugin manifest not found for package "${packageName}". `
        + `Expected: ${manifestPath}`,
    )
  }

  let manifest: PluginManifest
  try {
    const raw = await readFile(manifestPath, 'utf-8')
    manifest = validateManifest(JSON.parse(raw))
  } catch (error) {
    if (error instanceof DocxKitError) {
      throw error
    }
    throw new DocxKitError(
      'MANIFEST_INVALID',
      `Failed to parse manifest for "${packageName}"`,
      error,
    )
  }

  // Resolve the main entry relative to the package root
  const entryPath = `${pkgPath}/${manifest.main}`

  // Dynamic import the plugin module
  let mod: DocxPlugin | { default?: unknown }
  try {
    mod = (await import(entryPath)) as DocxPlugin | { default?: unknown }
  } catch (error) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Failed to import plugin module "${manifest.main}" from "${packageName}"`,
      error,
    )
  }

  // Extract the plugin — check for default export first, then named
  const exported = mod as { default?: unknown }
  const plugin = exported.default ?? mod

  if (!isDocxPlugin(plugin)) {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Module "${manifest.main}" from "${packageName}" does not export a valid DocxPlugin`,
    )
  }

  // If the manifest declares a different plugin name, verify it matches
  if (plugin.name !== manifest.plugin.name) {
    console.warn(
      `[docx-kit] Plugin name mismatch: manifest declares "${manifest.plugin.name}" `
        + `but plugin instance has "${plugin.name}"`,
    )
  }

  return { manifest, plugin: plugin as DocxPlugin }
}

// ---------- Helpers ----------

/**
 * Walk up from a resolved module path to the package root directory.
 *
 * Starting from a resolved module path (file or directory), walks
 * up the directory tree looking for a `package.json` file.
 *
 * @param resolvedPath - — The resolved module file or directory path
 * @returns The package root directory, or the original path if not found
 */
function resolvePackageRoot(resolvedPath: string): string {
  let dir = resolvedPath

  // If it's a file, start from its directory
  if (!dir.endsWith('/')) {
    const lastSlash = dir.lastIndexOf('/')
    if (lastSlash > 0) {
      dir = dir.slice(0, lastSlash)
    }
  }

  // Walk up until we find package.json
  for (let i = 0; i < 10; i++) {
    if (existsSync(`${dir}/package.json`)) {
      return dir
    }
    const parent = dir.slice(0, dir.lastIndexOf('/'))
    if (!parent || parent === dir) {
      break
    }
    dir = parent
  }

  // If no package.json found, return the original directory
  return resolvedPath
}
