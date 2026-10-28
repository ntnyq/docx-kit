import { DocxKitError as CoreError } from '@docxkit/core'
import { DocxKitError } from '@docxkit/errors'
import { describe, expect, it } from 'vitest'
import { createPreviewError, PREVIEW_ERROR_CODES } from '../src/errors'

describe('shared preview errors', () => {
  it('uses the same error class as the generation engine', () => {
    const cause = new Error('Network unavailable')
    const error = createPreviewError(
      PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED,
      'Preview fetch failed',
      cause,
    )

    expect(error).toBeInstanceOf(DocxKitError)
    expect(error).toBeInstanceOf(CoreError)
    expect(error.code).toBe(PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED)
    expect(error.cause).toBe(cause)
  })
})
