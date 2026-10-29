import { DocxKitError as CoreError } from '@docxkit/core'
import { DocxKitError, serializeDocxKitError } from '@docxkit/errors'
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
  it.each(Object.values(PREVIEW_ERROR_CODES))(
    'includes a documentation link for %s',
    code => {
      const error = createPreviewError(code, 'Original preview message')
      const docs = `https://docx-kit.ntnyq.dev/errors/docxkit_${code.toLowerCase()}/`
      expect(error.docs).toBe(docs)
      expect(serializeDocxKitError(error).docs).toBe(docs)
      expect(error.message).toBe('Original preview message')
    },
  )
})
