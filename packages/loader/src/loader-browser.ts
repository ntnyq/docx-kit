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
    const manifestUrl = new URL('docx-kit.plugin.json', moduleUrl)
    const response = await fetch(manifestUrl)
    if (!response.ok) {
      throw new DocxKitError(
        'MANIFEST_MISSING',
        `Plugin manifest could not be loaded from "${manifestUrl.href}": HTTP ${response.status}`,
      )
    }

    return load(moduleUrl, {
      authorizeManifest,
      manifest: await response.json(),
    })
  }
}

/** Create a browser plugin loader with URL and same-origin support. */
export function createPluginLoader(
  options: PluginLoaderOptions = {},
): PluginLoader {
  return new BrowserPluginLoader({
    ...options,
    kitVersion: options.kitVersion ?? DOCX_KIT_VERSION,
  })
}
