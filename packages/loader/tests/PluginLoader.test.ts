/* eslint-disable vitest/no-conditional-expect */
import {
  createPluginLoader,
  DocxKitError,
  ERROR_CODES,
  PluginLoader,
} from '@docxkit/core'
import { describe, expect, it, vi } from 'vitest'
import { calloutPlugin } from '../../../packages-plugins/callout/src/index'
import { watermarkPlugin } from '../../../packages-plugins/watermark/src/index'
import type { DocxPlugin, PluginManifest, PluginSource } from '@docxkit/core'

const callout = calloutPlugin() as DocxPlugin

// Reusable manifest template with all required fields
const TEST_MANIFEST: PluginManifest = {
  docxKit: '*',
  main: './dist/index.js',
  name: 'test',
  plugin: { name: 'test' },
  version: '1.0.0',
}

describe('PluginLoader', () => {
  describe('load() — inline source', () => {
    it('loads an inline plugin and returns null manifest', async () => {
      const loader = createPluginLoader()
      const source: PluginSource = { plugin: callout, type: 'inline' }

      const result = await loader.load(source)
      expect(result.plugin).toBe(callout)
      expect(result.manifest).toBeNull()
      expect(result.source).toBe(source)
    })

    it('preserves plugin name and type', async () => {
      const loader = createPluginLoader()
      const wm = watermarkPlugin() as DocxPlugin
      const result = await loader.load({ plugin: wm, type: 'inline' })
      expect(result.plugin.name).toBe('watermark')
    })
  })

  describe('load() — stub errors for platform-specific sources', () => {
    it('throws PLUGIN_LOAD_FAILED for npm source on base loader', async () => {
      const loader = createPluginLoader()
      try {
        await loader.load({ package: 'some-plugin', type: 'npm' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
        expect((err as DocxKitError).message).toContain('npm')
      }
    })

    it('throws PLUGIN_LOAD_FAILED for url source on base loader', async () => {
      const loader = createPluginLoader()
      try {
        await loader.load({ type: 'url', url: 'https://example.com/plugin.js' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
        expect((err as DocxKitError).message).toContain('browser')
      }
    })

    it('throws PLUGIN_LOAD_FAILED for local source on base loader', async () => {
      const loader = createPluginLoader()
      try {
        await loader.load({ path: './plugin.js', type: 'local' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
        expect((err as DocxKitError).message).toContain('Local')
      }
    })
  })

  describe('loadAll()', () => {
    it('collects successful loads and warns on failures', async () => {
      const loader = createPluginLoader()
      const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})

      const sources: PluginSource[] = [
        { plugin: callout, type: 'inline' },
        { package: 'nonexistent', type: 'npm' },
      ]

      const results = await loader.loadAll(sources)
      expect(results).toHaveLength(1)
      expect(results[0].plugin.name).toBe('callout')
      expect(warnSpy).toHaveBeenCalled()

      warnSpy.mockRestore()
    })

    it('returns all results when all sources succeed', async () => {
      const loader = createPluginLoader()
      const wm = watermarkPlugin() as DocxPlugin

      const sources: PluginSource[] = [
        { plugin: callout, type: 'inline' },
        { plugin: wm, type: 'inline' },
      ]

      const results = await loader.loadAll(sources)
      expect(results).toHaveLength(2)
      expect(results[0].plugin.name).toBe('callout')
      expect(results[1].plugin.name).toBe('watermark')
    })

    it('returns empty array when all sources fail', async () => {
      const loader = createPluginLoader()
      const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})

      const sources: PluginSource[] = [
        { package: 'nonexistent1', type: 'npm' },
        { type: 'url', url: 'https://bad.example.com/plugin.js' },
      ]

      const results = await loader.loadAll(sources)
      expect(results).toHaveLength(0)
      expect(warnSpy).toHaveBeenCalledTimes(2)

      warnSpy.mockRestore()
    })
  })

  describe('security policy — allowLoad', () => {
    it('blocks loading when allowLoad returns false', async () => {
      const loader = createPluginLoader({
        security: {
          allowLoad: () => false,
        },
      })

      try {
        await loader.load({ plugin: callout, type: 'inline' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
        expect((err as DocxKitError).message).toContain('security policy')
      }
    })

    it('allows loading when allowLoad returns true', async () => {
      const loader = createPluginLoader({
        security: {
          allowLoad: () => true,
        },
      })

      const result = await loader.load({ plugin: callout, type: 'inline' })
      expect(result.plugin.name).toBe('callout')
    })

    it('supports async allowLoad', async () => {
      const loader = createPluginLoader({
        security: {
          allowLoad: async () => false,
        },
      })

      try {
        await loader.load({ plugin: callout, type: 'inline' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
      }
    })
  })

  describe('security policy — allowExecute', () => {
    it('blocks execution when allowExecute returns false', async () => {
      const loader = createPluginLoader({
        security: {
          allowExecute: () => false,
        },
      })
      // Inject manifest to trigger allowExecute
      loader._forceManifestForTesting = TEST_MANIFEST

      try {
        await loader.load({ plugin: callout, type: 'inline' })
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
        expect((err as DocxKitError).message).toContain('security policy')
      }
    })

    it('allows execution when allowExecute returns true', async () => {
      const loader = createPluginLoader({
        security: {
          allowExecute: () => true,
        },
      })
      loader._forceManifestForTesting = TEST_MANIFEST

      const result = await loader.load({ plugin: callout, type: 'inline' })
      expect(result.plugin.name).toBe('callout')
    })
  })

  describe('version compatibility — _checkCompatibility', () => {
    it('accepts wildcard (*) range', () => {
      const loader = createPluginLoader({ kitVersion: '0.2.0' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '*' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      expect(() => testLoader._checkCompatibility(manifest)).not.toThrow()
    })

    it('accepts caret range with matching major version', () => {
      const loader = createPluginLoader({ kitVersion: '0.3.0' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '^0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      expect(() => testLoader._checkCompatibility(manifest)).not.toThrow()
    })

    it('rejects caret range with mismatched major version', () => {
      const loader = createPluginLoader({ kitVersion: '2.0.0' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '^0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      try {
        testLoader._checkCompatibility(manifest)
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_VERSION_MISMATCH,
        )
      }
    })

    it('accepts tilde range with matching major.minor', () => {
      const loader = createPluginLoader({ kitVersion: '0.2.5' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '~0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      expect(() => testLoader._checkCompatibility(manifest)).not.toThrow()
    })

    it('rejects tilde range with mismatched major.minor', () => {
      const loader = createPluginLoader({ kitVersion: '0.3.0' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '~0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      try {
        testLoader._checkCompatibility(manifest)
      } catch (err) {
        expect(err instanceof DocxKitError).toBe(true)
        expect((err as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_VERSION_MISMATCH,
        )
      }
    })
  })

  describe('constructor options', () => {
    it('uses default kitVersion when not provided', () => {
      const loader = createPluginLoader()
      expect(loader).toBeDefined()
    })

    it('uses provided kitVersion', () => {
      const loader = createPluginLoader({ kitVersion: '1.0.0' })
      expect(loader).toBeDefined()
    })

    it('defaults validateManifest to true', () => {
      const loader = createPluginLoader()
      expect(loader).toBeDefined()
    })

    it('allows disabling manifest validation', () => {
      const loader = createPluginLoader({ validateManifest: false })
      expect(loader).toBeDefined()
    })
  })
})

// Extend PluginLoader with a testing hook
// We add _forceManifestForTesting to override inline loading behavior
declare module '../../core/src/loader/PluginLoader' {
  interface PluginLoader {
    _forceManifestForTesting?: PluginManifest | null
  }
}

// Monkey-patch the _loadInline method on PluginLoader prototype
// to support testing the allowExecute security hook.
// @ts-expect-error — accessing protected member for test purposes
const originalLoadInline = PluginLoader.prototype._loadInline
// @ts-expect-error — assigning to protected member; return type widened for testing
PluginLoader.prototype._loadInline = function (
  this: PluginLoader & { _forceManifestForTesting?: PluginManifest | null },
  plugin: DocxPlugin,
): { manifest: PluginManifest | null; plugin: DocxPlugin } {
  if (this._forceManifestForTesting !== undefined) {
    return {
      manifest: this._forceManifestForTesting,
      plugin,
    }
  }
  return originalLoadInline.call(this, plugin)
}
