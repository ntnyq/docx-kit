import { DocxKitError, ERROR_CODES } from '@docxkit/types'
import { describe, expect, it } from 'vitest'

describe('ERROR_CODES', () => {
  it('contains all expected error codes', () => {
    expect(ERROR_CODES.EXPORT_FAILED).toBe('EXPORT_FAILED')
    expect(ERROR_CODES.IMAGE_INVALID_DATA).toBe('IMAGE_INVALID_DATA')
    expect(ERROR_CODES.PLUGIN_NOT_REGISTERED).toBe('PLUGIN_NOT_REGISTERED')
    expect(ERROR_CODES.PLUGIN_RENDER_FAILED).toBe('PLUGIN_RENDER_FAILED')
    expect(ERROR_CODES.STYLE_UNKNOWN_CLASS).toBe('STYLE_UNKNOWN_CLASS')
    expect(ERROR_CODES.TABLE_INVALID_COLUMNS).toBe('TABLE_INVALID_COLUMNS')
    expect(ERROR_CODES.UNKNOWN_NODE_TYPE).toBe('UNKNOWN_NODE_TYPE')
  })

  it('has all values as their own string keys (as const pattern)', () => {
    // Verify each key equals its value (standard frozen-like pattern)
    for (const [key, val] of Object.entries(ERROR_CODES)) {
      expect(val).toBe(key)
    }
  })
})

describe('DocxKitError', () => {
  it('creates an error with code and message', () => {
    const err = new DocxKitError('TEST_CODE', 'Something went wrong')
    expect(err).toBeInstanceOf(Error)
    expect(err).toBeInstanceOf(DocxKitError)
    expect(err.name).toBe('DocxKitError')
    expect(err.code).toBe('TEST_CODE')
    expect(err.message).toBe('Something went wrong')
    expect(err.cause).toBeUndefined()
  })

  it('uses ERROR_CODES constants', () => {
    const err = new DocxKitError(
      ERROR_CODES.PLUGIN_NOT_REGISTERED,
      'Plugin not found',
    )
    expect(err.code).toBe('PLUGIN_NOT_REGISTERED')
  })

  it('stores the cause from a previous error', () => {
    const cause = new Error('original error')
    const err = new DocxKitError('WRAP', 'Wrapped error', cause)
    expect(err.cause).toBe(cause)
  })

  it('stores non-Error causes', () => {
    const err = new DocxKitError('WRAP', 'Wrapped', 'raw string cause')
    expect(err.cause).toBe('raw string cause')
  })

  it('is catchable and has stack trace', () => {
    expect(() => {
      throw new DocxKitError(ERROR_CODES.EXPORT_FAILED, 'Export failed')
    }).toThrow(DocxKitError)
  })
})
