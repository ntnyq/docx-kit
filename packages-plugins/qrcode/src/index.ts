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

import {
  createImageRun,
  createPluginStyle,
  dataUrlToUint8Array,
  definePlugin,
} from '@docxkit/core'
import { Paragraph, TextRun } from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface QRCodeConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: QRCodeStyles
}

/**
 * Options for the QRCode plugin.
 */
export interface QRCodePluginOptions {
  /**
   * The text / URL to encode.
   */
  text: string
  /**
   * Optional caption text displayed below the QR code.
   */
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
  /**
   * QR code margin (in modules).
   * @default 1
   */
  margin?: number
  /**
   * QR code image size in pixels.
   * @default 128
   */
  size?: number
  /**
   * Per-part styles for this invocation.
   */
  styles?: QRCodeStyles
}

/**
 * Independently configurable QRCode parts.
 */
export interface QRCodeStyles {
  /**
   * Caption run typography.
   */
  caption?: TextStyleRule
  /**
   * Caption paragraph alignment, spacing and flow.
   */
  captionParagraph?: ParagraphStyleRule
  /**
   * Image paragraph alignment, spacing and flow.
   */
  imageParagraph?: ParagraphStyleRule
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
export function qrcodePlugin(pluginConfig: QRCodeConfig = {}) {
  return definePlugin<'qrcode', QRCodePluginOptions>({
    name: 'qrcode',
    async render(options, context) {
      const partStyle = (...parts: (keyof QRCodeStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const captionStyle = partStyle('caption')
      const captionParagraphStyle = partStyle('captionParagraph')
      const imageParagraphStyle = partStyle('imageParagraph')
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

      const paragraphs: Paragraph[] = [
        new Paragraph(imageParagraphStyle.paragraph({ children: [imageRun] })),
      ]

      if (options.caption) {
        paragraphs.push(
          new Paragraph(
            captionParagraphStyle.paragraph({
              children: [
                new TextRun(captionStyle.text({ text: options.caption })),
              ],
            }),
          ),
        )
      }

      return paragraphs
    },
  })
}
