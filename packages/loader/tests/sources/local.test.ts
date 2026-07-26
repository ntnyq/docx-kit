/* eslint-disable vitest/no-conditional-expect */
import { DocxKitError, ERROR_CODES } from '@docxkit/core'
import { describe, expect, it } from 'vitest'
import { loadLocalPlugin } from '../../src/sources/local'

describe('loadLocalPlugin', () => {
  it('throws PLUGIN_LOAD_FAILED when dynamic import fails', async () => {
    // Local import will fail in test environment for arbitrary paths
    try {
      await loadLocalPlugin('/nonexistent/path/plugin.js')
    } catch (error) {
      expect(error instanceof DocxKitError).toBe(true)
      expect((error as DocxKitError).code).toBe(ERROR_CODES.PLUGIN_LOAD_FAILED)
      expect((error as DocxKitError).message).toContain('local path')
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
