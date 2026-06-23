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

interface BufferConstructorLike {
  from(data: string, encoding: 'base64'): Uint8Array
}

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
  const payloadStart = dataUrl.indexOf(',')
  if (payloadStart === -1) {
    throw new Error('Expected a base64 data URL payload')
  }

  const base64 = dataUrl.slice(payloadStart + 1)

  // Browser path: use atob
  if (typeof atob === 'function') {
    const binary = atob(base64)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) {
      // eslint-disable-next-line unicorn/prefer-code-point -- base64 is always ASCII, charCodeAt is sufficient here
      bytes[i] = binary.charCodeAt(i)
    }
    return bytes
  }

  // Node.js path: use the global Buffer without importing `node:buffer`, so
  // neutral/browser bundles do not need to resolve Node built-ins.
  const bufferConstructor = Reflect.get(globalThis, 'Buffer') as unknown
  if (!isBufferConstructorLike(bufferConstructor)) {
    throw new Error('No base64 decoder is available in this environment')
  }
  return new Uint8Array(bufferConstructor.from(base64, 'base64'))
}

function isBufferConstructorLike(
  value: unknown,
): value is BufferConstructorLike {
  return (
    (typeof value === 'function'
      || (typeof value === 'object' && value !== null))
    && typeof Reflect.get(value, 'from') === 'function'
  )
}
