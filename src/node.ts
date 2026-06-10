/**
 * docx-kit Node.js platform entry.
 *
 * Re-exports all APIs available in Node.js environments, and documents
 * which APIs are **not** available and why.
 *
 * ## Available API overview
 *
 * | API | Status | Notes |
 * |-----|--------|-------|
 * | `saveDocument()` | ✅ Built-in | Writes .docx to disk via `node:fs/promises` |
 * | `dataUrlToUint8Array()` | ✅ Built-in | Decodes base64 via `node:buffer` |
 * | `echartsPlugin()` | ❌ Not available | Requires `window`/DOM for chart rendering |
 * | `normalizeImageData()` | ❌ Not built-in | `Blob` exists in Node 18+ but rarely used |
 *
 * For the full cross-platform builder API (`createDocx`, `defineStyles`, etc.),
 * import from `'docx-kit'` as usual.
 *
 * ## Usage
 *
 * ```ts
 * import { createDocx, defineStyles } from 'docx-kit'
 * import { saveDocument, dataUrlToUint8Array } from 'docx-kit/node'
 *
 * const doc = createDocx({ styles: defineStyles({ p: { fontSize: 12 } }) })
 *   .p('Hello from Node.js')
 *
 * await doc.save('output.docx')
 * // Equivalent to:
 * // const compiled = await doc.toDocument()
 * // await saveDocument(compiled, 'output.docx')
 * ```
 *
 * @module docx-kit/node
 * @packageDocumentation
 */

// ---------- Available APIs ----------

export { dataUrlToUint8Array, saveDocument } from './node/index'

// ---------- Unavailable APIs (documented only) ----------

/**
 * ### ❌ `echartsPlugin()` — Not available in Node.js
 *
 * The ECharts plugin requires a browser `window` and `document` to
 * render charts into a DOM container. Node.js has no native DOM.
 *
 * **Workarounds:**
 * - Use a server-side canvas library (e.g. `node-canvas` + `echarts`)
 * - Pre-render charts on the client and pass the image data to docx-kit
 *
 * @deprecated This API is not available in Node.js. See workarounds above.
 */
export declare const echartsPlugin: never

/**
 * ### ❌ `normalizeImageData()` — Not built-in for Node.js
 *
 * This utility converts `Blob` instances to `Uint8Array`. While `Blob`
 * is available in Node.js ≥ 18, it is rarely the carrier for image data
 * in Node workflows (most users work with `Buffer` or file paths directly).
 *
 * If you need Blob→Uint8Array in Node.js, use:
 * ```ts
 * const bytes = new Uint8Array(await blob.arrayBuffer())
 * ```
 *
 * @deprecated Use inline `new Uint8Array(await blob.arrayBuffer())` instead.
 */
export declare const normalizeImageData: never
