/**
 * Plugin search — npm keyword search via registry API.
 *
 * Searches the npm registry for packages tagged with
 * `keyword:docx-kit-plugin` and resolves their manifests.
 *
 * @module registry/PluginSearch
 */

import { validateManifest } from '../loader/manifest'
import type { PluginManifest } from '../loader/manifest'
import type {
  NpmSearchResult,
  QualityScore,
  RegistryPluginEntry,
} from './types'

/** npm registry search API endpoint. */
const NPM_SEARCH_URL = 'https://registry.npmjs.org/-/v1/search'

/**
 * Keyword used to tag docx-kit plugins on npm.
 */
const PLUGIN_KEYWORD = 'docx-kit-plugin'

/**
 * Get a specific plugin from the npm registry by package name.
 *
 * Fetches the full package metadata and resolves its manifest.
 *
 * @param packageName - — npm package name (e.g. `docx-kit-plugin-chart`)
 * @returns Registry entry or null if not found / not a docx-kit plugin
 */
export async function getPlugin(
  packageName: string,
): Promise<RegistryPluginEntry | null> {
  try {
    const url = `https://registry.npmjs.org/${encodeURIComponent(packageName)}`
    const response = await fetch(url)

    if (!response.ok) {
      return null
    }

    const data = (await response.json()) as {
      'dist-tags': { latest: string }
      name: string
      versions: Record<string, { keywords?: string[] }>
      description?: string
      keywords?: string[]
      repository?: { url: string }
    }

    // Verify this is actually a docx-kit plugin
    const latestVersion = data['dist-tags'].latest
    const versionData = data.versions[latestVersion]
    const keywords = versionData?.keywords ?? data.keywords ?? []

    if (!keywords.includes(PLUGIN_KEYWORD)) {
      return null
    }

    const manifest = extractManifest(data, latestVersion)
    if (!manifest) {
      return null
    }

    return {
      description: data.description ?? '',
      downloads: 0,
      manifest,
      name: data.name,
      quality: computeQuality(data, latestVersion),
      version: latestVersion,
    }
  } catch {
    return null
  }
}

/**
 * Search the npm registry for docx-kit plugins.
 *
 * Uses the npm `/-/v1/search` endpoint with the
 * `keyword:docx-kit-plugin` filter plus any user query.
 *
 * @param query - — Additional search terms (optional)
 * @returns Array of matching registry entries
 */
export async function searchPlugins(
  query?: string,
): Promise<RegistryPluginEntry[]> {
  const text = query
    ? `keyword:${PLUGIN_KEYWORD} ${query}`
    : `keyword:${PLUGIN_KEYWORD}`

  const url = `${NPM_SEARCH_URL}?text=${encodeURIComponent(text)}&size=25`

  const response = await fetch(url)

  if (!response.ok) {
    throw new Error(
      `npm registry search failed: ${response.status} ${response.statusText}`,
    )
  }

  const data = (await response.json()) as {
    objects: NpmSearchResult[]
    total: number
  }

  return data.objects.map(resolveEntry)
}

/**
 * Compute a quality score for a plugin based on available signals.
 */
function computeQuality(
  data: {
    versions: Record<string, { keywords?: string[] }>
    keywords?: string[]
    repository?: { url: string }
  },
  version: string,
): QualityScore {
  const versionData = data.versions[version]
  const keywords = versionData?.keywords ?? data.keywords ?? []

  return {
    hasManifest: true,
    hasTests: keywords.includes('vitest') || keywords.includes('jest'),
    hasTypescript: keywords.includes('typescript'),
    stars: 0, // Would need GitHub API to resolve
  }
}

/**
 * Extract and validate a manifest from npm package metadata.
 *
 * Looks for a `docx-kit.plugin.json` in the package files
 * or falls back to a generated manifest from package metadata.
 */
function extractManifest(
  data: {
    name: string
    versions: Record<string, { keywords?: string[] }>
    keywords?: string[]
  },
  version: string,
): PluginManifest | null {
  // Try to construct a manifest from package metadata
  const shortName = data.name.replace(/^docx-kit-plugin-/, '')

  try {
    return validateManifest({
      docxKit: '*',
      main: './dist/index.js',
      name: data.name,
      plugin: { name: shortName },
      version,
    })
  } catch {
    return null
  }
}

/**
 * Resolve a npm search result into a registry entry.
 */
function resolveEntry(result: NpmSearchResult): RegistryPluginEntry {
  const pkg = result.package

  const manifest: PluginManifest = {
    docxKit: '*',
    main: './dist/index.js',
    name: pkg.name,
    version: pkg.version,
    plugin: {
      name: pkg.name.replace(/^docx-kit-plugin-/, ''),
    },
  }

  return {
    description: pkg.description ?? '',
    downloads: result.downloads ?? 0,
    manifest,
    name: pkg.name,
    version: pkg.version,
    quality: {
      hasManifest: true,
      stars: 0,
      hasTests:
        (pkg.keywords?.includes('vitest') ?? false)
        || (pkg.keywords?.includes('jest') ?? false),
      hasTypescript:
        (pkg.keywords?.includes('typescript') ?? false)
        || pkg.name.includes('ts'),
    },
  }
}
