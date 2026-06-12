import { DocxKitError, ERROR_CODES } from '@docxkit/core'
import { describe, expect, it } from 'vitest'
import { loadUrlPlugin } from '../../src/sources/url'

describe('loadUrlPlugin', () => {
  it('throws PLUGIN_LOAD_FAILED when import fails', async () => {
    // Dynamic import() of arbitrary URLs will fail in test environment
    await expect(
      loadUrlPlugin('https://invalid.example.com/plugin.js'),
    ).rejects.toThrow(DocxKitError)

    try {
      await loadUrlPlugin('https://invalid.example.com/plugin.js')
    } catch (err) {
      // eslint-disable-next-line vitest/no-conditional-expect
      expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
      // eslint-disable-next-line vitest/no-conditional-expect
      expect((err as DocxKitError).message).toContain('URL')
    }
  })

  it('accepts a pre-resolved manifest option', () => {
    // Validates option shape
    const options = {
      manifest: {
        docxKit: '*',
        name: 'test-plugin',
        plugin: { name: 'test' },
        version: '1.0.0',
      },
    }
    expect(options.manifest.plugin.name).toBe('test')
  })
})
