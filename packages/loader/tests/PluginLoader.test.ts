/* eslint-disable vitest/no-conditional-expect */
import { fileURLToPath } from 'node:url'
import {
  createPluginLoader,
  DocxKitError,
  ERROR_CODES,
  PluginLoader,
} from '@docxkit/core'
import { describe, expect, it, vi } from 'vitest'
import { calloutPlugin } from '../../../packages-plugins/callout/src/index'
import { watermarkPlugin } from '../../../packages-plugins/watermark/src/index'
import { createPluginLoader as createBrowserPluginLoader } from '../src/loader-browser'
import { createPluginLoader as createNodePluginLoader } from '../src/loader-node'
import type {
  DocxPlugin,
  PluginManifest,
  PluginManifestAuthorizer,
  PluginSource,
} from '@docxkit/core'

const callout = calloutPlugin() as DocxPlugin

// Reusable manifest template with all required fields
const TEST_MANIFEST: PluginManifest = {
  docxKit: '*',
  main: './dist/index.js',
  name: 'test',
  plugin: { name: 'test' },
  version: '1.0.0',
}

class AuthorizedNpmLoader extends PluginLoader {
  readonly events: string[] = []

  protected override async _loadNpm(
    _packageName: string,
    authorizeManifest: PluginManifestAuthorizer,
  ) {
    this.events.push('manifest')
    const manifest = await authorizeManifest(TEST_MANIFEST)
    this.events.push('execute')
    return { manifest, plugin: callout }
  }
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
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_LOAD_FAILED,
        )
        expect((error as DocxKitError).message).toContain('npm')
      }
    })

    it('throws PLUGIN_LOAD_FAILED for url source on base loader', async () => {
      const loader = createPluginLoader()
      try {
        await loader.load({ type: 'url', url: 'https://example.com/plugin.js' })
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_LOAD_FAILED,
        )
        expect((error as DocxKitError).message).toContain('browser')
      }
    })

    it('throws PLUGIN_LOAD_FAILED for local source on base loader', async () => {
      const loader = createPluginLoader()
      try {
        await loader.load({ path: './plugin.js', type: 'local' })
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_LOAD_FAILED,
        )
        expect((error as DocxKitError).message).toContain('Local')
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
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_LOAD_FAILED,
        )
        expect((error as DocxKitError).message).toContain('security policy')
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
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_LOAD_FAILED,
        )
      }
    })
  })

  describe('security policy — allowExecute', () => {
    it('blocks execution when allowExecute returns false', async () => {
      const loader = new AuthorizedNpmLoader({
        kitVersion: '1.0.0',
        security: {
          allowExecute: () => {
            loader.events.push('policy')
            return false
          },
        },
      })

      await expect(
        loader.load({ package: 'test-plugin', type: 'npm' }),
      ).rejects.toMatchObject({
        code: ERROR_CODES.PLUGIN_LOAD_FAILED,
      })
      expect(loader.events).toEqual(['manifest', 'policy'])
    })

    it('allows execution when allowExecute returns true', async () => {
      const loader = new AuthorizedNpmLoader({
        kitVersion: '1.0.0',
        security: {
          allowExecute: () => {
            loader.events.push('policy')
            return true
          },
        },
      })

      const result = await loader.load({
        package: 'test-plugin',
        type: 'npm',
      })
      expect(result.plugin.name).toBe('callout')
      expect(loader.events).toEqual(['manifest', 'policy', 'execute'])
    })
  })

  describe('Node platform factory', () => {
    it('loads a local plugin using its adjacent manifest', async () => {
      const loader = createNodePluginLoader()
      const pluginDirectory = fileURLToPath(
        new URL('./fixtures/node-plugin', import.meta.url),
      )

      const result = await loader.load({
        path: pluginDirectory,
        type: 'local',
      })

      expect(result.plugin.name).toBe('fixture')
      expect(result.manifest?.main).toBe('./index.mjs')
    })
  })

  describe('browser platform factory', () => {
    it('loads a URL plugin after fetching its adjacent manifest', async () => {
      const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
        Response.json(
          {
            ...TEST_MANIFEST,
            docxKit: '^0.4.0',
            main: './index.mjs',
            plugin: { name: 'fixture' },
          },
          { status: 200 },
        ),
      )
      const loader = createBrowserPluginLoader()
      const moduleUrl = new URL(
        './fixtures/node-plugin/index.mjs',
        import.meta.url,
      ).href

      const result = await loader.load({ type: 'url', url: moduleUrl })

      expect(result.plugin.name).toBe('fixture')
      expect(fetchSpy).toHaveBeenCalledWith(
        new URL('./fixtures/node-plugin/docx-kit.plugin.json', import.meta.url),
      )
      fetchSpy.mockRestore()
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

    it('accepts caret range within the same zero-major minor line', () => {
      const loader = createPluginLoader({ kitVersion: '0.2.5' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '^0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      expect(() => testLoader._checkCompatibility(manifest)).not.toThrow()
    })

    it('rejects caret range outside the zero-major minor line', () => {
      const loader = createPluginLoader({ kitVersion: '0.3.0' })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit: '^0.2.0' }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      try {
        testLoader._checkCompatibility(manifest)
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
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
      } catch (error) {
        expect(error instanceof DocxKitError).toBe(true)
        expect((error as DocxKitError).code).toBe(
          ERROR_CODES.PLUGIN_VERSION_MISMATCH,
        )
      }
    })

    it.each([
      ['0.3.0', '0.3.0'],
      ['0.3.4', '>=0.3.0 <0.4.0'],
      ['0.4.0-beta.1', '^0.4.0-beta.1'],
    ])('accepts version %s for range %s', (kitVersion, docxKit) => {
      const loader = createPluginLoader({ kitVersion })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }
      expect(() => testLoader._checkCompatibility(manifest)).not.toThrow()
    })

    it.each([
      ['0.3.1', '0.3.0'],
      ['0.4.0-beta.1', '^0.3.0'],
      ['not-a-version', '*'],
      ['0.3.0', 'not-a-range'],
    ])('rejects version %s for range %s', (kitVersion, docxKit) => {
      const loader = createPluginLoader({ kitVersion })
      const manifest: PluginManifest = { ...TEST_MANIFEST, docxKit }
      const testLoader = loader as unknown as {
        _checkCompatibility(m: PluginManifest): void
      }

      expect(() => testLoader._checkCompatibility(manifest)).toThrowError(
        DocxKitError,
      )
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
