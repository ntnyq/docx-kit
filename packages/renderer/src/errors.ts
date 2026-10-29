/**
 * `@docxkit/renderer` — Error codes and factory.
 *
 * Renderer-local error codes extend the shared error model from
 * `@docxkit/errors` without depending on the document generation engine.
 *
 * All errors are `DocxKitError` instances so consumers can use a single
 * error class across the entire docx-kit ecosystem.
 *
 * @module renderer/errors
 */

import { createDocxKitError, ERROR_CODES } from '@docxkit/errors'
import type { DocxKitError } from '@docxkit/errors'

/**
 * Renderer-specific error codes.
 *
 * - `PREVIEW_INPUT_INVALID` — Input is null, undefined, or not a recognized type
 * - `PREVIEW_FETCH_FAILED` — Failed to fetch DOCX from remote URL
 * - `PREVIEW_RENDER_FAILED` — `docx-preview` `renderAsync` threw an error
 * - `MICROSOFT_URL_REQUIRED` — Microsoft renderer requires a URL string input
 */
export const PREVIEW_ERROR_CODES = {
  MICROSOFT_URL_REQUIRED: ERROR_CODES.MICROSOFT_URL_REQUIRED,
  PREVIEW_FETCH_FAILED: ERROR_CODES.PREVIEW_FETCH_FAILED,
  PREVIEW_INPUT_INVALID: ERROR_CODES.PREVIEW_INPUT_INVALID,
  PREVIEW_RENDER_FAILED: ERROR_CODES.PREVIEW_RENDER_FAILED,
} as const

/**
 * Union of all renderer error codes.
 */
export type PreviewErrorCode =
  (typeof PREVIEW_ERROR_CODES)[keyof typeof PREVIEW_ERROR_CODES]

/**
 * Create a renderer-specific {@link DocxKitError}.
 *
 * @param code - The preview error code
 * @param message - Human-readable error message
 * @param cause - The underlying error (if any)
 * @returns A DocxKitError instance with the preview code prefix
 */
export function createPreviewError(
  code: PreviewErrorCode,
  message: string,
  cause?: unknown,
): DocxKitError {
  return createDocxKitError(code, { detail: message }, { cause })
}
