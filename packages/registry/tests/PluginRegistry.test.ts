import { afterEach, describe, expect, it, vi } from 'vitest'
import { createPluginRegistry, PluginRegistry } from '../src/PluginRegistry'
import { getPlugin, searchPlugins } from '../src/PluginSearch'
import type { RegistryPluginEntry } from '../src/types'

// Mock the search functions
vi.mock('../src/PluginSearch', () => ({
  getPlugin: vi.fn(),
  searchPlugins: vi.fn(),
}))

const mockSearchPlugins = vi.mocked(searchPlugins)
const mockGetPlugin = vi.mocked(getPlugin)

const MOCK_ENTRY: RegistryPluginEntry = {
  description: 'A chart plugin for docx-kit',
  downloads: 150,
  name: 'docx-kit-plugin-chart',
  version: '1.0.0',
  manifest: {
    docxKit: '^0.2.0',
    main: './dist/index.js',
    name: 'docx-kit-plugin-chart',
    plugin: { name: 'chart' },
    version: '1.0.0',
  },
  quality: {
    hasManifest: true,
    hasTests: true,
    hasTypescript: true,
    stars: 5,
  },
}

describe('PluginRegistry', () => {
  afterEach(() => {
    vi.clearAllMocks()
  })

  describe('search()', () => {
    it('returns matching plugins from npm registry', async () => {
      mockSearchPlugins.mockResolvedValueOnce([MOCK_ENTRY])

      const registry = new PluginRegistry()
      const results = await registry.search('chart')

      expect(results).toHaveLength(1)
      expect(results[0].name).toBe('docx-kit-plugin-chart')
      expect(mockSearchPlugins).toHaveBeenCalledWith('chart')
    })

    it('returns empty array when no results found', async () => {
      mockSearchPlugins.mockResolvedValueOnce([])

      const registry = new PluginRegistry()
      const results = await registry.search('nonexistent')

      expect(results).toHaveLength(0)
    })

    it('caches results within TTL', async () => {
      mockSearchPlugins.mockResolvedValue([MOCK_ENTRY])

      const registry = new PluginRegistry({ cacheTTL: 60000 })
      const first = await registry.search('chart')
      const second = await registry.search('chart')

      expect(first).toHaveLength(1)
      expect(second).toHaveLength(1)
      // Should only call the search function once (second call uses cache)
      expect(mockSearchPlugins).toHaveBeenCalledTimes(1)
    })

    it('refreshes cache after TTL expires', async () => {
      mockSearchPlugins.mockResolvedValue([MOCK_ENTRY])

      const registry = new PluginRegistry({ cacheTTL: 0 })
      const first = await registry.search('chart')
      const second = await registry.search('chart')

      expect(first).toHaveLength(1)
      expect(second).toHaveLength(1)
      // TTL is 0, so cache expires immediately — two calls
      expect(mockSearchPlugins).toHaveBeenCalledTimes(2)
    })
  })

  describe('get()', () => {
    it('returns specific plugin from npm registry', async () => {
      mockGetPlugin.mockResolvedValueOnce(MOCK_ENTRY)

      const registry = new PluginRegistry()
      const result = await registry.get('docx-kit-plugin-chart')

      expect(result).not.toBeNull()
      expect(result!.name).toBe('docx-kit-plugin-chart')
      expect(mockGetPlugin).toHaveBeenCalledWith('docx-kit-plugin-chart')
    })

    it('returns null for unknown package', async () => {
      mockGetPlugin.mockResolvedValueOnce(null)

      const registry = new PluginRegistry()
      const result = await registry.get('nonexistent-plugin')

      expect(result).toBeNull()
    })

    it('caches get results within TTL', async () => {
      mockGetPlugin.mockResolvedValue(MOCK_ENTRY)

      const registry = new PluginRegistry({ cacheTTL: 60000 })
      const first = await registry.get('docx-kit-plugin-chart')
      const second = await registry.get('docx-kit-plugin-chart')

      expect(first).not.toBeNull()
      expect(second).not.toBeNull()
      expect(mockGetPlugin).toHaveBeenCalledTimes(1)
    })
  })

  describe('refresh()', () => {
    it('does not let a pending search repopulate the refreshed cache', async () => {
      const pending = createDeferred<RegistryPluginEntry[]>()
      const fresh = { ...MOCK_ENTRY, version: '2.0.0' }
      mockSearchPlugins
        .mockReturnValueOnce(pending.promise)
        .mockResolvedValueOnce([fresh])
      const registry = new PluginRegistry()
      const stale = registry.search('chart')
      registry.refresh()
      pending.resolve([MOCK_ENTRY])
      await stale
      expect(await registry.search('chart')).toEqual([fresh])
      expect(await registry.search('chart')).toEqual([fresh])
      expect(mockSearchPlugins).toHaveBeenCalledTimes(2)
    })

    it('does not let a pending package lookup repopulate the refreshed cache', async () => {
      const pending = createDeferred<RegistryPluginEntry | null>()
      const fresh = { ...MOCK_ENTRY, version: '2.0.0' }
      mockGetPlugin
        .mockReturnValueOnce(pending.promise)
        .mockResolvedValueOnce(fresh)
      const registry = new PluginRegistry()
      const stale = registry.get(MOCK_ENTRY.name)
      registry.refresh()
      pending.resolve(MOCK_ENTRY)
      await stale
      expect(await registry.get(MOCK_ENTRY.name)).toEqual(fresh)
      expect(await registry.get(MOCK_ENTRY.name)).toEqual(fresh)
      expect(mockGetPlugin).toHaveBeenCalledTimes(2)
    })

    it('clears cache and forces fresh data', async () => {
      mockSearchPlugins.mockResolvedValue([MOCK_ENTRY])

      const registry = new PluginRegistry({ cacheTTL: 60000 })
      await registry.search('chart')

      // Refresh clears cache
      registry.refresh()

      await registry.search('chart')
      expect(mockSearchPlugins).toHaveBeenCalledTimes(2)
    })
  })

  describe('createPluginRegistry()', () => {
    it('creates a registry instance', () => {
      const registry = createPluginRegistry()
      expect(registry).toBeInstanceOf(PluginRegistry)
    })

    it('passes options to the registry', () => {
      const registry = createPluginRegistry({ cacheTTL: 10000 })
      expect(registry).toBeInstanceOf(PluginRegistry)
    })
  })
})

function createDeferred<T>() {
  let resolve: (value: T) => void = () => {}
  const promise = new Promise<T>(resolvePromise => {
    resolve = resolvePromise
  })
  return {
    promise,
    resolve,
  }
}
