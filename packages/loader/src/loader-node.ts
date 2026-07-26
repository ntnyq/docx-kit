/**
 * @docxkit/loader/node — Node.js platform plugin loader.
 *
 * @packageDocumentation
 */

import { readFile, stat } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { PluginLoader } from '@docxkit/core'
import { loadLocalPlugin } from './sources/local'
import { loadNpmPlugin } from './sources/npm'
import { DOCX_KIT_VERSION } from './version'
import type {
  PluginLoaderOptions,
  PluginManifestAuthorizer,
} from '@docxkit/core'

export * from './index'
export { loadNpmPlugin } from './sources/npm'

class NodePluginLoader extends PluginLoader {
  protected override async _loadLocal(
    sourcePath: string,
    authorizeManifest: PluginManifestAuthorizer,
  ) {
    const resolvedPath = sourcePath.startsWith('file:')
      ? fileURLToPath(sourcePath)
      : path.resolve(sourcePath)
    const sourceStats = await stat(resolvedPath)
    const packageRoot = sourceStats.isDirectory()
      ? resolvedPath
      : path.dirname(resolvedPath)
    const manifestPath = path.join(packageRoot, 'docx-kit.plugin.json')
    const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as unknown
    const entry =
      typeof manifest === 'object'
      && manifest !== null
      && 'main' in manifest
      && typeof manifest.main === 'string'
        ? path.resolve(packageRoot, manifest.main)
        : resolvedPath

    return loadLocalPlugin(pathToFileURL(entry).href, {
      authorizeManifest,
      manifest,
    })
  }

  protected override async _loadNpm(
    packageName: string,
    authorizeManifest: PluginManifestAuthorizer,
  ) {
    const result = await loadNpmPlugin(packageName, { authorizeManifest })
    if (!result.manifest) {
      throw new Error('npm plugin loader returned no manifest')
    }
    return { manifest: result.manifest, plugin: result.plugin }
  }
}

/** Create a Node.js plugin loader with npm and local-file support. */
export function createPluginLoader(
  options: PluginLoaderOptions = {},
): PluginLoader {
  return new NodePluginLoader({
    ...options,
    kitVersion: options.kitVersion ?? DOCX_KIT_VERSION,
  })
}
