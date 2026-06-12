import { DocxKitError, validateManifest } from '@docxkit/core'
import { describe, expect, it } from 'vitest'
import type { PluginManifest } from '@docxkit/core'

const VALID_MANIFEST: PluginManifest = {
  docxKit: '^0.2.0',
  main: './dist/index.js',
  name: 'my-chart-plugin',
  version: '1.0.0',
  plugin: {
    author: { email: 'john@example.com', name: 'John Doe' },
    description: 'A chart plugin for docx-kit',
    license: 'MIT',
    name: 'myChart',
  },
}

describe('validateManifest', () => {
  describe('valid manifests', () => {
    it('accepts a manifest with all required fields', () => {
      const result = validateManifest(VALID_MANIFEST)
      expect(result).toEqual(VALID_MANIFEST)
    })

    it('accepts a manifest with minimal required fields', () => {
      const minimal = {
        docxKit: '*',
        name: 'my-plugin',
        plugin: { name: 'myPlugin' },
        version: '0.1.0',
      }
      const result = validateManifest(minimal)
      expect(result.name).toBe('my-plugin')
      expect(result.version).toBe('0.1.0')
      expect(result.docxKit).toBe('*')
      expect(result.plugin.name).toBe('myPlugin')
    })

    it('accepts a manifest with pre-release version', () => {
      const manifest = { ...VALID_MANIFEST, version: '1.0.0-beta.1' }
      expect(validateManifest(manifest).version).toBe('1.0.0-beta.1')
    })

    it('accepts a manifest with build metadata', () => {
      const manifest = { ...VALID_MANIFEST, version: '1.0.0+build.123' }
      expect(validateManifest(manifest).version).toBe('1.0.0+build.123')
    })

    it('accepts complex semver ranges for docxKit', () => {
      const ranges = ['>=0.5.0 <1.0.0', '^0.2.0', '~0.2.0', '*', '0.2.0']
      for (const range of ranges) {
        const manifest = { ...VALID_MANIFEST, docxKit: range }
        expect(validateManifest(manifest).docxKit).toBe(range)
      }
    })
  })

  describe('invalid inputs', () => {
    it('rejects null input', () => {
      expect(() =>
        validateManifest(null as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects non-object input (string)', () => {
      expect(() =>
        validateManifest('not an object' as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects non-object input (number)', () => {
      expect(() =>
        validateManifest(42 as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects array input', () => {
      expect(() =>
        validateManifest([] as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects undefined input', () => {
      expect(() =>
        validateManifest(undefined as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })
  })

  describe('required field validation', () => {
    function expectManifestError(
      manifest: unknown,
      checker: (err: DocxKitError) => void,
    ): void {
      try {
        validateManifest(manifest as unknown as Record<string, unknown>)
      } catch (err) {
        if (err instanceof DocxKitError) {
          checker(err)
          return
        }
        throw err
      }
      throw new Error('Expected validateManifest to throw')
    }

    it('rejects missing name', () => {
      const manifest = { ...VALID_MANIFEST, name: undefined }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
      expectManifestError(manifest, err => {
        expect(err.code).toBe('MANIFEST_INVALID')
        expect(err.message).toContain('manifest.name')
      })
    })

    it('rejects empty string name', () => {
      const manifest = { ...VALID_MANIFEST, name: '' }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects missing version', () => {
      const manifest = { ...VALID_MANIFEST, version: undefined }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
      expectManifestError(manifest, err => {
        expect(err.message).toContain('manifest.version')
      })
    })

    it('rejects empty string version', () => {
      const manifest = { ...VALID_MANIFEST, version: '' }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects missing docxKit', () => {
      const manifest = { ...VALID_MANIFEST, docxKit: undefined }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects empty string docxKit', () => {
      const manifest = { ...VALID_MANIFEST, docxKit: '' }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects missing plugin object', () => {
      const manifest = { ...VALID_MANIFEST, plugin: undefined }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects null plugin object', () => {
      const manifest = { ...VALID_MANIFEST, plugin: null }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects array plugin object', () => {
      const manifest = { ...VALID_MANIFEST, plugin: [] }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects missing plugin.name', () => {
      const manifest = { ...VALID_MANIFEST, plugin: { description: 'test' } }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
      expectManifestError(manifest, err => {
        expect(err.message).toContain('manifest.plugin.name')
      })
    })

    it('rejects empty string plugin.name', () => {
      const manifest = { ...VALID_MANIFEST, plugin: { name: '' } }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })
  })

  describe('semver validation', () => {
    it('rejects invalid semver for version', () => {
      const invalidVersions = ['1.0', '1', 'abc', 'v1.0.0', '1.0.0.0']
      for (const v of invalidVersions) {
        const manifest = { ...VALID_MANIFEST, version: v }
        expect(() =>
          validateManifest(manifest as unknown as Record<string, unknown>),
        ).toThrow(DocxKitError)
      }
    })
  })

  describe('range validation', () => {
    it('rejects invalid range for docxKit', () => {
      const invalidRanges = ['abc', '1.0.0a', '!invalid']
      for (const r of invalidRanges) {
        const manifest = { ...VALID_MANIFEST, docxKit: r }
        expect(() =>
          validateManifest(manifest as unknown as Record<string, unknown>),
        ).toThrow(DocxKitError)
      }
    })
  })

  describe('optional field validation', () => {
    it('accepts manifest without main field', () => {
      const manifest = { ...VALID_MANIFEST, main: undefined }
      expect(
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toBeDefined()
    })

    it('rejects non-string main field', () => {
      const manifest = { ...VALID_MANIFEST, main: 42 }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('accepts manifest without types field', () => {
      const manifest = { ...VALID_MANIFEST, types: undefined }
      expect(
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toBeDefined()
    })

    it('rejects non-string types field', () => {
      const manifest = { ...VALID_MANIFEST, types: 123 }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('accepts valid dependencies object', () => {
      const manifest = { ...VALID_MANIFEST, dependencies: { docx: '^9.7.1' } }
      expect(
        validateManifest(manifest as unknown as Record<string, unknown>)
          .dependencies,
      ).toEqual({ docx: '^9.7.1' })
    })

    it('rejects non-object dependencies', () => {
      const manifest = { ...VALID_MANIFEST, dependencies: 'invalid' }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('rejects null dependencies', () => {
      const manifest = { ...VALID_MANIFEST, dependencies: null }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })

    it('accepts valid peerDependencies object', () => {
      const manifest = {
        ...VALID_MANIFEST,
        peerDependencies: { echarts: '^5.0.0' },
      }
      expect(
        validateManifest(manifest as unknown as Record<string, unknown>)
          .peerDependencies,
      ).toEqual({ echarts: '^5.0.0' })
    })

    it('rejects non-object peerDependencies', () => {
      const manifest = { ...VALID_MANIFEST, peerDependencies: 'invalid' }
      expect(() =>
        validateManifest(manifest as unknown as Record<string, unknown>),
      ).toThrow(DocxKitError)
    })
  })
})
