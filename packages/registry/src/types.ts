/**
 * Plugin registry types.
 *
 * @module registry/types
 */

import type { PluginManifest } from '@docxkit/core'

/**
 * Search result from the npm registry API.
 */
export interface NpmSearchResult {
  /** Download count (if available). */
  downloads?: number

  /** Package name. */
  package: {
    description: string
    name: string
    version: string
    keywords?: string[]
    links?: { repository?: string }
  }

  /** Score object from npm search. */
  score?: {
    detail: { popularity: number; quality: number }
    final: number
  }
}

/**
 * Quality score for a registry plugin.
 *
 * Based on observable signals from the npm package and
 * its repository (if linked).
 */
export interface QualityScore {
  /** Whether the package includes a docx-kit plugin manifest. */
  hasManifest: boolean

  /** Whether the package has test files detected. */
  hasTests: boolean

  /** Whether the package includes TypeScript source/types. */
  hasTypescript: boolean

  /** GitHub star count (0 if no repo linked). */
  stars: number
}

/**
 * A plugin entry in the registry.
 *
 * Represents metadata about a published docx-kit plugin,
 * including its npm package details, manifest, and quality score.
 */
export interface RegistryPluginEntry {
  /** Package description from npm. */
  description: string

  /** Weekly download count from npm. */
  downloads: number

  /** Parsed plugin manifest. */
  manifest: PluginManifest

  /** npm package name. */
  name: string

  /** Quality assessment based on available signals. */
  quality: QualityScore

  /** Latest published version. */
  version: string
}
