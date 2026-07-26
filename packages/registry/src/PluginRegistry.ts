/**
 * Plugin registry — local registry class for discovering and
 * querying docx-kit plugins.
 *
 * @module registry/PluginRegistry
 */

import { getPlugin, searchPlugins } from './PluginSearch'
import type { RegistryPluginEntry } from './types'

/**
 * Cache entry with TTL.
 */
interface CacheEntry {
  entries: RegistryPluginEntry[]
  expiresAt: number
}

/**
 * Default cache TTL in milliseconds (5 minutes).
 */
const DEFAULT_CACHE_TTL = 5 * 60 * 1000

/**
 * Plugin registry — discovers and queries docx-kit plugins
 * from the npm registry with local caching.
 *
 * @example
 * ```ts
 * import { PluginRegistry } from 'docx-kit/registry'
 *
 * const registry = new PluginRegistry()
 * const plugins = await registry.search('chart')
 * console.log(plugins.map(p => p.name))
 *
 * const chartPlugin = await registry.get('docx-kit-plugin-chart')
 * if (chartPlugin?.manifest) {
 *   console.log(chartPlugin.manifest.plugin.name)
 * }
 * ```
 */
export class PluginRegistry {
  private cache = new Map<string, CacheEntry>()
  private readonly cacheTTL: number

  constructor(options?: { cacheTTL?: number }) {
    this.cacheTTL = options?.cacheTTL ?? DEFAULT_CACHE_TTL
  }

  /**
   * Get a specific plugin by its npm package name.
   *
   * @param packageName - — npm package name (e.g. `docx-kit-plugin-chart`)
   * @returns Registry entry or null if not found
   */
  async get(packageName: string): Promise<RegistryPluginEntry | null> {
    const cacheKey = `get:${packageName}`

    const cached = this.cache.get(cacheKey)
    if (cached && cached.expiresAt > Date.now()) {
      return cached.entries[0] ?? null
    }

    const entry = await getPlugin(packageName)

    if (entry) {
      this.cache.set(cacheKey, {
        entries: [entry],
        expiresAt: Date.now() + this.cacheTTL,
      })
    }

    return entry
  }

  /**
   * Refresh the cache — removes all cached entries.
   *
   * Subsequent `search()` or `get()` calls will fetch
   * fresh data from the npm registry.
   */
  refresh(): void {
    this.cache.clear()
  }

  /**
   * Search for plugins matching a query.
   *
   * Results are cached for the configured TTL. Subsequent
   * calls with the same query return cached results until
   * the cache expires.
   *
   * @param query - — Search terms (optional; returns all plugins if empty)
   * @returns Array of matching registry entries
   */
  async search(query?: string): Promise<RegistryPluginEntry[]> {
    const cacheKey = `search:${query ?? ''}`

    const cached = this.cache.get(cacheKey)
    if (cached && cached.expiresAt > Date.now()) {
      return cached.entries
    }

    const entries = await searchPlugins(query)

    this.cache.set(cacheKey, {
      entries,
      expiresAt: Date.now() + this.cacheTTL,
    })

    return entries
  }
}

/**
 * Create a PluginRegistry instance.
 *
 * @param options - — Registry configuration
 * @param options.cacheTTL - — Cache TTL in milliseconds
 * @returns A configured registry
 */
export function createPluginRegistry(options?: {
  cacheTTL?: number
}): PluginRegistry {
  return new PluginRegistry(options)
}
