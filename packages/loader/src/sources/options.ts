import { DocxKitError, validateManifest } from '@docxkit/core'
import type { PluginManifest, PluginManifestAuthorizer } from '@docxkit/core'

export interface ExternalPluginLoadOptions {
  /**
   * Manifest authorization callback supplied by PluginLoader.
   *
   * When present, the source loader must invoke it before importing code.
   */
  authorizeManifest?: PluginManifestAuthorizer
  /**
   * Manifest resolved independently from the executable module.
   */
  manifest?: unknown
}

export async function resolveManifest(
  options?: ExternalPluginLoadOptions,
): Promise<PluginManifest | null> {
  if (options?.manifest === undefined) {
    if (options?.authorizeManifest) {
      throw new DocxKitError(
        'MANIFEST_MISSING',
        'A plugin manifest is required before external code can be authorized',
      )
    }
    return null
  }

  if (options.authorizeManifest) {
    return await options.authorizeManifest(options.manifest)
  }

  return validateManifest(options.manifest)
}
