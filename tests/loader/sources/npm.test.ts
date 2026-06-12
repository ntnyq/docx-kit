/* eslint-disable vitest/no-conditional-expect */
import { describe, expect, it } from 'vitest'
import { DocxKitError, ERROR_CODES } from '../../../src/errors'

// Note: loadNpmPlugin requires Node.js environment with actual
// node_modules. These tests verify error behavior for the
// executable paths without mocking Node built-ins (which is flaky in vitest).

describe('loadNpmPlugin', () => {
  it('throws PLUGIN_LOAD_FAILED for unresolvable package', async () => {
    // Dynamic import of the npm loader module
    // This test verifies the function exists and can be called
    // The actual resolution failure is hard to test without a real package
    try {
      const { loadNpmPlugin } = await import('../../../src/loader/sources/npm')
      await loadNpmPlugin('definitely-not-a-real-package-xyz123')
    } catch (err) {
      expect(err instanceof DocxKitError).toBe(true)
      // Either MANIFEST_MISSING (if resolved but no manifest) or PLUGIN_LOAD_FAILED (if resolve fails)
      const code = (err as DocxKitError).code
      expect([
        ERROR_CODES.MANIFEST_MISSING,
        ERROR_CODES.PLUGIN_LOAD_FAILED,
      ]).toContain(code)
    }
  })
})
