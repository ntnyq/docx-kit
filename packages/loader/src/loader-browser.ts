/**
 * `@docxkit/loader/browser` — Browser platform plugin loader.
 *
 * @packageDocumentation
 */

import { DocxKitError, PluginLoader } from '@docxkit/core'
import { loadLocalPlugin } from './sources/local'
import { loadUrlPlugin } from './sources/url'
import { DOCX_KIT_VERSION } from './version'
import type {
  PluginLoaderOptions,
  PluginManifestAuthorizer,
} from '@docxkit/core'

export * from './index'
export { loadUrlPlugin } from './sources/url'

class BrowserPluginLoader extends PluginLoader {
  protected override async _loadLocal(
    path: string,
    authorizeManifest: PluginManifestAuthorizer,
  ) {
    const url = new URL(path, globalThis.location?.href)
    return this.loadBrowserSource(url.href, authorizeManifest, loadLocalPlugin)
  }

  protected override async _loadUrl(
    url: string,
    authorizeManifest: PluginManifestAuthorizer,
  ) {
    return this.loadBrowserSource(url, authorizeManifest, loadUrlPlugin)
  }

  private async loadBrowserSource(
    moduleUrl: string,
    authorizeManifest: PluginManifestAuthorizer,
    load: typeof loadUrlPlugin,
  ) {
    const sourceUrl = new URL(moduleUrl, globalThis.location?.href)
    const isManifest = sourceUrl.pathname.endsWith('/docx-kit.plugin.json')
    const manifestUrl = isManifest
      ? sourceUrl
      : new URL('docx-kit.plugin.json', sourceUrl)
    const response = await fetch(manifestUrl)
    if (!response.ok) {
      throw new DocxKitError(
        'MANIFEST_MISSING',
        `Plugin manifest could not be loaded from "${manifestUrl.href}": HTTP ${response.status}`,
      )
    }

    const manifest = await authorizeManifest(await response.json())
    const entryUrl =
      isManifest || sourceUrl.pathname.endsWith('/')
        ? new URL(manifest.main, manifestUrl).href
        : sourceUrl.href
    const packageUrl = new URL('./', manifestUrl)
    const entry = new URL(entryUrl)
    if (
      entry.origin !== packageUrl.origin
      || !entry.pathname.startsWith(packageUrl.pathname)
    ) {
      throw new DocxKitError(
        'MANIFEST_INVALID',
        `Plugin entry "${manifest.main}" resolves outside its package URL`,
      )
    }

    return load(entryUrl, {
      manifest,
      authorizeManifest: () => manifest,
    })
  }
}

/**
 * Create a browser plugin loader with URL and same-origin support.
 *
 * @param options - Optional compatibility, manifest validation, and security settings
 * @returns A browser plugin loader using the supplied settings
 */
export function createPluginLoader(
  options: PluginLoaderOptions = {},
): PluginLoader {
  return new BrowserPluginLoader({
    ...options,
    kitVersion: options.kitVersion ?? DOCX_KIT_VERSION,
  })
}
