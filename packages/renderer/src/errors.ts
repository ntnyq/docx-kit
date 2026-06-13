/**
 * @docxkit/renderer — Error codes and factory.
 *
 * Renderer-local error codes. Cannot extend `@docxkit/core`'s `ERROR_CODES`
 * because the dependency direction is renderer → core, never the reverse.
 *
 * All errors are `DocxKitError` instances so consumers can use a single
 * error class across the entire docx-kit ecosystem.
 *
 * @module renderer/errors
 */

import { DocxKitError } from '@docxkit/core'

/**
 * Renderer-specific error codes.
 *
 * - `PREVIEW_INPUT_INVALID` — Input is null, undefined, or not a recognized type
 * - `PREVIEW_FETCH_FAILED` — Failed to fetch DOCX from remote URL
 * - `PREVIEW_RENDER_FAILED` — `docx-preview` `renderAsync` threw an error
 * - `MICROSOFT_URL_REQUIRED` — Microsoft renderer requires a URL string input
 */
export const PREVIEW_ERROR_CODES = {
  MICROSOFT_URL_REQUIRED: 'MICROSOFT_URL_REQUIRED',
  PREVIEW_FETCH_FAILED: 'PREVIEW_FETCH_FAILED',
  PREVIEW_INPUT_INVALID: 'PREVIEW_INPUT_INVALID',
  PREVIEW_RENDER_FAILED: 'PREVIEW_RENDER_FAILED',
} as const

/** Union of all renderer error codes. */
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
  return new DocxKitError(code, message, cause)
}
