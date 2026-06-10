/**
 * Node.js base64 data-URL decoder (uses `Buffer` from `node:buffer`).
 *
 * @module node/dataUrl
 */

/**
 * Decode a base64 data-URL to raw bytes using Node.js `Buffer`.
 *
 * Strips the `"data:*;base64,"` prefix and decodes via
 * `Buffer.from(base64, 'base64')`.
 *
 * @param dataUrl - — A base64 data-URI string (e.g. `"data:image/png;base64,iVBO..."`)
 * @returns Raw bytes as `Uint8Array`
 *
 * @example
 * ```ts
 * import { dataUrlToUint8Array } from 'docx-kit/node'
 *
 * const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw...')
 * ```
 */
export async function dataUrlToUint8Array(
  dataUrl: string,
): Promise<Uint8Array> {
  const { Buffer } = await import('node:buffer')
  const base64 = dataUrl.split(',')[1]
  return new Uint8Array(Buffer.from(base64, 'base64'))
}
