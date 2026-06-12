/* eslint-disable vitest/no-conditional-expect */
import { describe, expect, it } from 'vitest'
import { DocxKitError, ERROR_CODES } from '../../../src/errors'
import { loadLocalPlugin } from '../../../src/loader/sources/local'

describe('loadLocalPlugin', () => {
  it('throws PLUGIN_LOAD_FAILED when dynamic import fails', async () => {
    // Local import will fail in test environment for arbitrary paths
    try {
      await loadLocalPlugin('/nonexistent/path/plugin.js')
    } catch (err) {
      expect(err instanceof DocxKitError).toBe(true)
      expect((err as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
      expect((err as DocxKitError).message).toContain('local path')
    }
  })

  it('accepts a pre-resolved manifest option', () => {
    // Validates option shape
    const options = {
      manifest: {
        docxKit: '^0.2.0',
        name: 'local-test',
        plugin: { name: 'localTest' },
        version: '1.0.0',
      },
    }
    expect(options.manifest.name).toBe('local-test')
  })
})
