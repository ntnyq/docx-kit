/**
 * Cross-platform base64 and percent-encoded data-URL decoder.
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
 * Decode a data URL to raw bytes (works in both browser & Node.js).
 *
 * Supports base64 payloads and percent-encoded bytes, including SVG data URLs.
 *
 * @param dataUrl - — A data-URI string (e.g. `"data:image/png;base64,iVBO..."`)
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
  if (!dataUrl.startsWith('data:') || payloadStart === -1) {
    throw new Error('Expected a data URL payload')
  }

  const payload = dataUrl.slice(payloadStart + 1)
  if (!/;base64$/i.test(dataUrl.slice(0, payloadStart))) {
    // Decode bytes directly: percent escapes can contain non-UTF-8 image data.
    const chunks = payload.split(/(%[\da-f]{2})/i)
    const bytes = chunks.flatMap(chunk =>
      /^%[\da-f]{2}$/i.test(chunk)
        ? [Number.parseInt(chunk.slice(1), 16)]
        : [...new TextEncoder().encode(chunk)],
    )
    return Uint8Array.from(bytes)
  }
  const base64 = decodeURIComponent(payload)

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
