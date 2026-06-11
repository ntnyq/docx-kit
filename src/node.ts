/**
 * docx-kit Node.js platform entry — import from `'docx-kit/node'`.
 *
 * Re-exports all shared APIs plus Node.js–specific utilities
 * (filesystem save, Buffer-based base64 decoding).
 *
 * ## Usage
 *
 * ```ts
 * import { createDocx, defineStyles } from 'docx-kit/node'
 * import { saveDocument } from 'docx-kit/node'
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

// ---------- Shared API ----------
export * from './shared'

// ---------- Node.js–specific APIs ----------

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
