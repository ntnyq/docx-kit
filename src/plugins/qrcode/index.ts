/**
 * QR Code plugin — embeds a QR code image in the document.
 *
 * Uses the `qrcode` package (peer dependency) to generate
 * a QR code PNG and renders it as an inline `ImageRun`.
 *
 * @module plugins/qrcode
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(qrcodePlugin())
 *   .h1('Scan Me')
 *   .plugin('qrcode', { text: 'https://example.com', caption: 'Visit us' })
 *   .save('qrcode.docx')
 * ```
 */

import { Paragraph } from 'docx'
import { definePlugin } from '../../types/plugin'
import { dataUrlToUint8Array } from '../../utils/dataUrl'
import { createImageRun } from '../../utils/image'

/**
 * Options for the QRCode plugin.
 */
export interface QRCodePluginOptions {
  /** The text / URL to encode. */
  text: string
  /** Optional caption text displayed below the QR code. */
  caption?: string
  /**
   * QR error correction level.
   *
   * - `"L"` — ~7% recovery
   * - `"M"` — ~15% recovery
   * - `"Q"` — ~25% recovery
   * - `"H"` — ~30% recovery
   *
   * @default "M"
   */
  errorCorrectionLevel?: 'H' | 'L' | 'M' | 'Q'
  /** QR code margin (in modules). @default 1 */
  margin?: number
  /** QR code image size in pixels. @default 128 */
  size?: number
}

/**
 * Create a QRCode plugin instance.
 *
 * The plugin dynamically imports the `qrcode` package at render time,
 * keeping it an optional peer dependency of docx-kit core.
 *
 * @returns A configured DocxPlugin for `'qrcode'`
 *
 * @example
 * ```ts
 * import { createDocx, qrcodePlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(qrcodePlugin())
 *   .plugin('qrcode', {
 *     text: 'https://example.com',
 *     size: 256,
 *     errorCorrectionLevel: 'H',
 *     caption: 'Scan to visit',
 *   })
 * ```
 */
export function qrcodePlugin() {
  return definePlugin<'qrcode', QRCodePluginOptions>({
    name: 'qrcode',
    async render(options) {
      // Dynamic import: peer dep `qrcode` is not bundled in core.
      const QRCode =
        (await import('qrcode')).default ?? (await import('qrcode'))

      const size = options.size ?? 128
      const dataUrl = await QRCode.toDataURL(options.text, {
        errorCorrectionLevel: options.errorCorrectionLevel ?? 'M',
        margin: options.margin ?? 1,
        width: size,
      })

      const imageData = await dataUrlToUint8Array(dataUrl)

      const imageRun = createImageRun({
        data: imageData,
        height: size,
        type: 'png',
        width: size,
      })

      const paragraphs: Paragraph[] = [new Paragraph({ children: [imageRun] })]

      if (options.caption) {
        paragraphs.push(new Paragraph({ text: options.caption }))
      }

      return paragraphs
    },
  })
}
