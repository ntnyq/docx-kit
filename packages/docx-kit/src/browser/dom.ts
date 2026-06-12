/**
 * Browser DOM utilities — base64 decoding and Blob/image handling.
 *
 * @module browser/dom
 */

export { dataUrlToUint8Array } from '@docxkit/core'

/**
 * Normalize image data for `ImageRun` — converts `Blob` to `Uint8Array`.
 *
 * In the browser, image data often arrives as a `Blob` (e.g. from
 * `<input type="file">` or `fetch()`). `ImageRun` expects raw bytes,
 * so we convert via `Blob.arrayBuffer()`.
 *
 * @param data - — The raw image data (Blob, Uint8Array, ArrayBuffer, string)
 * @returns Normalized data suitable for `ImageRun`
 *
 * @example
 * ```ts
 * import { normalizeImageData } from 'docx-kit/browser'
 *
 * const file = document.querySelector('input[type=file]').files[0]
 * const data = await normalizeImageData(file) // Blob → Uint8Array
 * ```
 */
export async function normalizeImageData(
  data: unknown,
): Promise<string | ArrayBuffer | Uint8Array> {
  if (typeof Blob !== 'undefined' && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer())
  }
  return data as string | ArrayBuffer | Uint8Array
}
