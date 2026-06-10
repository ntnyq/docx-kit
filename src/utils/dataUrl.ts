/**
 * Cross-platform base64 data-URL decoder.
 *
 * Auto-detects the runtime environment and uses the appropriate
 * implementation:
 * - **Node.js**: `Buffer.from(base64, 'base64')`
 * - **Browser**: `atob(base64)` + manual byte population
 *
 * This is the internal shared version used by the compiler and
 * built-in plugins. For platform-specific direct access, import
 * from `'docx-kit/node'` or `'docx-kit/browser'`.
 *
 * @module utils/dataUrl
 */

/**
 * Decode a base64 data-URL to raw bytes (works in both browser & Node.js).
 *
 * Strips the `"data:*;base64,"` prefix and decodes using the appropriate
 * runtime API.
 *
 * @param dataUrl - — A base64 data-URI string (e.g. `"data:image/png;base64,iVBO..."`)
 * @returns Raw bytes as `Uint8Array`
 *
 * @example
 * ```ts
 * import { dataUrlToUint8Array } from 'docx-kit'
 *
 * const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw...')
 * ```
 */
export async function dataUrlToUint8Array(
  dataUrl: string,
): Promise<Uint8Array> {
  const base64 = dataUrl.split(',')[1]

  // Browser path: use atob
  if (typeof atob === 'function') {
    const binary = atob(base64)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.codePointAt(i) ?? 0
    }
    return bytes
  }

  // Node.js path: use Buffer
  const { Buffer } = await import('node:buffer')
  return new Uint8Array(Buffer.from(base64, 'base64'))
}
