/**
 * Browser DOM utilities — base64 decoding and Blob/image handling.
 *
 * @module browser/dom
 */

/**
 * Decode a base64 data-URL to raw bytes using the browser `atob` API.
 *
 * Strips the `"data:*;base64,"` prefix, decodes the base64 string
 * via `atob()`, and populates a `Uint8Array` from each character's
 * code point.
 *
 * @param dataUrl - — A base64 data-URI string (e.g. `"data:image/png;base64,iVBO..."`)
 * @returns Raw bytes as `Uint8Array`
 *
 * @example
 * ```ts
 * import { dataUrlToUint8Array } from 'docx-kit/browser'
 *
 * const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw...')
 * ```
 */
export async function dataUrlToUint8Array(
  dataUrl: string,
): Promise<Uint8Array> {
  const base64 = dataUrl.split(',')[1]
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    // eslint-disable-next-line unicorn/prefer-code-point -- base64 is always ASCII, charCodeAt is sufficient here
    bytes[i] = binary.charCodeAt(i)
  }
  return bytes
}

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
