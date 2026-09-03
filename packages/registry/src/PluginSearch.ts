/**
 * Plugin search — npm keyword search via registry API.
 *
 * Searches the npm registry for packages tagged with
 * `keyword:docx-kit-plugin` and resolves their manifests.
 *
 * @module registry/PluginSearch
 */

import { validateManifest } from '@docxkit/core'
import type { PluginManifest } from '@docxkit/core'
import type {
  NpmSearchResult,
  QualityScore,
  RegistryPluginEntry,
} from './types'

/**
 * npm registry search API endpoint.
 */
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

    const manifest = await fetchManifest(data.name, latestVersion)

    return {
      description: data.description ?? '',
      downloads: 0,
      manifest,
      name: data.name,
      quality: computeQuality(data, latestVersion, manifest !== null),
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
 *
 * Checks for TypeScript and test keywords in the package metadata.
 *
 * @param data - — npm package metadata
 * @param data.versions - — Package version map
 * @param data.keywords - — Top-level package keywords
 * @param data.repository - — Repository info
 * @param data.repository.url - — Repository URL
 * @param version - — The package version to inspect
 * @param hasManifest - — Whether a published manifest was fetched and validated
 * @returns Quality score object
 */
function computeQuality(
  data: {
    versions: Record<string, { keywords?: string[] }>
    keywords?: string[]
    repository?: { url: string }
  },
  version: string,
  hasManifest: boolean,
): QualityScore {
  const versionData = data.versions[version]
  const keywords = versionData?.keywords ?? data.keywords ?? []

  return {
    hasManifest,
    hasTests: keywords.includes('vitest') || keywords.includes('jest'),
    hasTypescript: keywords.includes('typescript'),
    stars: 0, // Would need GitHub API to resolve
  }
}

/**
 * Fetch and validate a published manifest.
 *
 * @param packageName - — npm package name
 * @param version - — The package version
 * @returns Validated manifest or null if it is missing or invalid
 */
async function fetchManifest(
  packageName: string,
  version: string,
): Promise<PluginManifest | null> {
  try {
    const url = `https://unpkg.com/${packageName}@${version}/docx-kit.plugin.json`
    const response = await fetch(url)

    if (!response.ok) {
      return null
    }

    return validateManifest(await response.json())
  } catch {
    return null
  }
}

/**
 * Resolve a npm search result into a registry entry.
 *
 * Constructs a RegistryPluginEntry from npm's search metadata. Search results
 * do not prove that a package contains a valid manifest, so the manifest is
 * left unverified until the package is fetched directly.
 *
 * @param result - — Raw npm search result
 * @returns A fully populated registry plugin entry
 */
function resolveEntry(result: NpmSearchResult): RegistryPluginEntry {
  const pkg = result.package

  return {
    description: pkg.description ?? '',
    downloads: result.downloads ?? 0,
    manifest: null,
    name: pkg.name,
    version: pkg.version,
    quality: {
      hasManifest: false,
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
