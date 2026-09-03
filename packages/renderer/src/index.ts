/**
 * `@docxkit/renderer` — Browser DOCX preview for docx-kit.
 *
 * Renders `.docx` files directly in the browser using either
 * [`docx-preview`](https://github.com/VolodymyrBaydalka/docxjs) (DOM-based,
 * default) or Microsoft Office Online (iframe, opt-in).
 *
 * ## Install
 *
 * ```sh
 * npm install @docxkit/renderer
 * ```
 *
 * ## Quick start
 *
 * ```ts
 * import { createDocxPreview } from '@docxkit/renderer'
 *
 * const preview = createDocxPreview(document.getElementById('app')!)
 *
 * // From a Blob
 * const blob = await doc.toBlob()
 * await preview.render(blob)
 *
 * // From a URL
 * await preview.render('https://example.com/document.docx')
 *
 * // Clean up
 * preview.destroy()
 * ```
 *
 * @packageDocumentation
 */

export { createDocxPreview } from './preview'
export { PREVIEW_ERROR_CODES } from './errors'
export type { PreviewErrorCode } from './errors'
export type {
  DocxInput,
  DocxPreview,
  DocxPreviewOptions,
  RendererKind,
} from './types'
