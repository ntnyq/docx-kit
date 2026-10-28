import { DocxKitError, ERROR_CODES } from '@docxkit/errors'
import { describe, expect, it } from 'vitest'
import {
  ERROR_CODES as CORE_ERROR_CODES,
  DocxKitError as CoreError,
  defineStyles,
} from '../src/index'

describe('shared error exports', () => {
  it('re-exports the shared error class and codes without duplicating them', () => {
    expect(CoreError).toBe(DocxKitError)
    expect(CORE_ERROR_CODES).toBe(ERROR_CODES)
  })

  it('reports invalid style inheritance with the shared error class', () => {
    expect(() => defineStyles({ title: { extends: 'missing' } })).toThrow(
      DocxKitError,
    )
    expect(() => defineStyles({ title: { extends: ['missing'] } })).toThrow(
      expect.objectContaining({ code: ERROR_CODES.STYLE_UNKNOWN_CLASS }),
    )
  })
})
