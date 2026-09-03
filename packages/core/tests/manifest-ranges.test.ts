import { describe, expect, it } from 'vitest'
import { validateManifest } from '../src/loader/manifest'

describe('manifest semver ranges', () => {
  it.each([
    '^0.4.0 || ^1.0.0',
    '^0.4.0-beta.1',
    '^0.4.0+build.123',
    '0.4.x',
    '0.4.0 - 0.5.0',
    '>=0.4.0 <1.0.0',
    '  ~0.4.0  ',
    '*',
  ])('accepts %s', docxKit => {
    expect(
      validateManifest({
        docxKit,
        main: './dist/index.js',
        name: 'example',
        plugin: { name: 'example' },
        version: '1.0.0',
      }).docxKit,
    ).toBe(docxKit)
  })

  it.each(['not-a-range', '^', '1.0.0a', '^0.4.0 || invalid', '>=0.4.0 <'])(
    'rejects invalid range %s with MANIFEST_INVALID',
    docxKit => {
      expect(() =>
        validateManifest({
          docxKit,
          main: './dist/index.js',
          name: 'example',
          plugin: { name: 'example' },
          version: '1.0.0',
        }),
      ).toThrowError(
        expect.objectContaining({
          code: 'MANIFEST_INVALID',
          message: `Plugin manifest field "manifest.docxKit" is not a valid semver range: "${docxKit}"`,
          name: 'DocxKitError',
        }),
      )
    },
  )
})
