/* eslint-disable vitest/no-conditional-expect */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { getPlugin, searchPlugins } from '../src/PluginSearch'

// Mock global fetch
const mockFetch = vi.fn()
vi.stubGlobal('fetch', mockFetch)

describe('searchPlugins', () => {
  afterEach(() => {
    mockFetch.mockReset()
  })

  it('searches npm registry with keyword filter', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        total: 1,
        objects: [
          {
            score: { final: 0.8 },
            package: {
              description: 'A chart plugin',
              keywords: ['docx-kit-plugin', 'chart'],
              name: 'docx-kit-plugin-chart',
              version: '1.0.0',
            },
          },
        ],
      }),
    })

    const results = await searchPlugins('chart')

    expect(results).toHaveLength(1)
    expect(results[0].name).toBe('docx-kit-plugin-chart')
    expect(results[0].manifest.plugin.name).toBe('chart')

    // Verify URL contains keyword and query (URL-encoded)
    const calledUrl = mockFetch.mock.calls[0][0] as string
    expect(calledUrl).toContain('keyword%3Adocx-kit-plugin')
    expect(calledUrl).toContain('chart')
  })

  it('returns empty array when no results', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({ objects: [], total: 0 }),
    })

    const results = await searchPlugins('nonexistent')
    expect(results).toHaveLength(0)
  })

  it('throws on registry error', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 500,
      statusText: 'Internal Server Error',
    })

    try {
      await searchPlugins('test')
    } catch (error) {
      expect(error instanceof Error).toBe(true)
      expect((error as Error).message).toContain('500')
    }
  })
})

describe('getPlugin', () => {
  afterEach(() => {
    mockFetch.mockReset()
  })

  it('returns plugin entry for valid package', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        description: 'A chart plugin',
        'dist-tags': { latest: '1.0.0' },
        keywords: ['docx-kit-plugin'],
        name: 'docx-kit-plugin-chart',
        versions: {
          '1.0.0': { keywords: ['docx-kit-plugin', 'chart'] },
        },
      }),
    })

    const result = await getPlugin('docx-kit-plugin-chart')

    expect(result).not.toBeNull()
    expect(result!.name).toBe('docx-kit-plugin-chart')
    expect(result!.manifest.plugin.name).toBe('chart')
  })

  it('returns null for non-existent package', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 404,
      statusText: 'Not Found',
    })

    const result = await getPlugin('nonexistent-plugin')
    expect(result).toBeNull()
  })

  it('returns null for package without docx-kit-plugin keyword', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        description: 'Some random package',
        'dist-tags': { latest: '1.0.0' },
        keywords: ['random', 'other'],
        name: 'random-package',
        versions: {
          '1.0.0': { keywords: ['random', 'other'] },
        },
      }),
    })

    const result = await getPlugin('random-package')
    expect(result).toBeNull()
  })

  it('returns null on fetch error', async () => {
    mockFetch.mockRejectedValueOnce(new Error('Network error'))

    const result = await getPlugin('docx-kit-plugin-chart')
    expect(result).toBeNull()
  })
})
