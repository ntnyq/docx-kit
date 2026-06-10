/**
 * docx-kit Browser platform entry.
 *
 * Re-exports all APIs available in browser environments, and documents
 * which APIs are **not** available and why.
 *
 * ## Available API overview
 *
 * | API | Status | Notes |
 * |-----|--------|-------|
 * | `dataUrlToUint8Array()` | ✅ Built-in | Decodes base64 via `atob()` |
 * | `normalizeImageData()` | ✅ Built-in | Converts `Blob` → `Uint8Array` |
 * | `saveDocument()` | ❌ Not available | No filesystem access in browsers |
 * | `echartsPlugin()` | ✅ Built-in | Renders charts in the DOM (import from `'docx-kit'`) |
 *
 * For the full cross-platform builder API (`createDocx`, `defineStyles`, etc.),
 * import from `'docx-kit'` as usual.
 *
 * ## Usage
 *
 * ```ts
 * import { createDocx, defineStyles, echartsPlugin } from 'docx-kit'
 * import { dataUrlToUint8Array, normalizeImageData } from 'docx-kit/browser'
 *
 * const doc = createDocx({ styles: defineStyles({ p: { fontSize: 12 } }) })
 *   .use(echartsPlugin())
 *   .p('Hello from the browser')
 *
 * // Use .toBlob() + URL.createObjectURL() for download
 * const blob = await doc.toBlob()
 * const url = URL.createObjectURL(blob)
 * const a = document.createElement('a')
 * a.href = url
 * a.download = 'output.docx'
 * a.click()
 * ```
 *
 * @module docx-kit/browser
 * @packageDocumentation
 */

// ---------- Available APIs ----------

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
