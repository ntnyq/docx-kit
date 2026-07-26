/**
 * @docxkit/renderer — Input normalization.
 *
 * Converts any {@link DocxInput} to a `Blob` so the rest of the
 * pipeline can work with a single type.
 *
 * @module renderer/normalize
 */

import { createPreviewError, PREVIEW_ERROR_CODES } from './errors'
import type { DocxInput } from './types'

/** Standard DOCX MIME type. */
export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/** Result of normalizing a {@link DocxInput} to a Blob. */
export interface NormalizedDocx {
  /** The DOCX data as a Blob. */
  blob: Blob
  /**
   * An object URL created via `URL.createObjectURL`, or `null` if none was
   * created. The caller is responsible for revoking this URL via
   * `URL.revokeObjectURL` when no longer needed.
   */
  objectUrl: string | null
}

/**
 * Convert any {@link DocxInput} to a Blob.
 *
 * Conversion rules:
 * - `Blob | File` — used as-is
 * - `ArrayBuffer` — wrapped in a new Blob with the DOCX MIME type
 * - `Uint8Array` — wrapped in a new Blob with the DOCX MIME type
 * - `string` (URL) — fetched via `fetch()`, response body returned as Blob
 * - anything else — throws `PREVIEW_INPUT_INVALID`
 *
 * @param input - The DOCX input to normalize
 * @param signal - Optional `AbortSignal` to cancel fetch operations
 * @returns A `NormalizedDocx` containing the Blob and an optional object URL
 * @throws {DocxKitError} `PREVIEW_INPUT_INVALID` for invalid inputs
 * @throws {DocxKitError} `PREVIEW_FETCH_FAILED` for fetch errors
 */
export async function normalizeDocxInput(
  input: DocxInput,
  signal?: AbortSignal,
): Promise<NormalizedDocx> {
  // Blob or File — pass through unchanged
  if (input instanceof Blob) {
    return { blob: input, objectUrl: null }
  }

  // ArrayBuffer — wrap in a Blob
  if (input instanceof ArrayBuffer) {
    return {
      blob: new Blob([input], { type: DOCX_MIME }),
      objectUrl: null,
    }
  }

  // Uint8Array — wrap in a Blob
  if (input instanceof Uint8Array) {
    return {
      blob: new Blob([input as BlobPart], { type: DOCX_MIME }),
      objectUrl: null,
    }
  }

  // String URL — fetch
  if (typeof input === 'string') {
    return await fetchDocxFromUrl(input, signal)
  }

  // Invalid input
  const actualType =
    input === null ? 'null' : Array.isArray(input) ? 'array' : typeof input
  throw createPreviewError(
    PREVIEW_ERROR_CODES.PREVIEW_INPUT_INVALID,
    `Expected Blob, File, ArrayBuffer, Uint8Array, or URL string, got ${actualType}`,
  )
}

/**
 * Fetch a DOCX file from a URL and return it as a Blob.
 *
 * @internal
 */
async function fetchDocxFromUrl(
  url: string,
  signal?: AbortSignal,
): Promise<NormalizedDocx> {
  try {
    const response = await fetch(url, { signal })
    if (!response.ok) {
      throw createPreviewError(
        PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED,
        `Failed to fetch DOCX from URL: HTTP ${response.status} ${response.statusText}`,
      )
    }
    const blob = await response.blob()
    return { blob, objectUrl: null }
  } catch (error) {
    // Re-throw AbortError as-is so callers can distinguish cancellation
    if (error instanceof Error && error.name === 'AbortError') {
      throw error
    }
    // Re-throw existing DocxKitError
    if (error instanceof Error && error.message.startsWith('[')) {
      throw error
    }
    throw createPreviewError(
      PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED,
      `Failed to fetch DOCX from URL: ${(error as Error).message ?? String(error)}`,
      error,
    )
  }
}
