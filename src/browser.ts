/**
 * docx-kit Browser platform entry — the default import target (`'docx-kit'`).
 *
 * Re-exports all shared APIs plus browser-specific utilities.
 * For Node.js–only APIs (filesystem write), import from `'docx-kit/node'`.
 *
 * ## Usage
 *
 * ```ts
 * import { createDocx, defineStyles, echartsPlugin } from 'docx-kit'
 * import { normalizeImageData } from 'docx-kit'
 *
 * const doc = createDocx({ styles: defineStyles({ p: { fontSize: 12 } }) })
 *   .use(echartsPlugin())
 *   .p('Hello from the browser')
 *
 * // Download via Blob
 * const blob = await doc.toBlob()
 * const url = URL.createObjectURL(blob)
 * const a = document.createElement('a')
 * a.href = url
 * a.download = 'output.docx'
 * a.click()
 * URL.revokeObjectURL(url)
 * ```
 *
 * @module docx-kit
 * @packageDocumentation
 */

// ---------- Shared API ----------
export * from './shared'

// ---------- Browser-specific APIs ----------

export { dataUrlToUint8Array, normalizeImageData } from './browser/index'

// ---------- Unavailable APIs (documented only) ----------

/**
 * ### ❌ `saveDocument()` — Not available in browsers
 *
 * Browsers do not provide direct filesystem access for security reasons.
 * To trigger a file download in the browser, use `doc.toBlob()` with
 * `URL.createObjectURL()` and a temporary `<a>` element:
 *
 * ```ts
 * const blob = await doc.toBlob()
 * const url = URL.createObjectURL(blob)
 * const a = document.createElement('a')
 * a.href = url
 * a.download = 'output.docx'
 * a.click()
 * URL.revokeObjectURL(url)
 * ```
 *
 * @deprecated This API is not available in browsers. Use `doc.toBlob()` instead.
 */
export declare const saveDocument: never
