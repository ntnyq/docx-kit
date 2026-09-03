/**
 * Watermark plugin — text watermark rendered as a styled paragraph.
 *
 * Renders a large, semi-transparent (gray) text paragraph that can be
 * placed in a section header or footer to simulate a watermark.
 *
 * @module plugins/watermark
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(watermarkPlugin())
 *   .plugin('watermark', { text: 'CONFIDENTIAL', color: 'FF0000' })
 *   .save('watermark.docx')
 * ```
 */

import { definePlugin } from '@docxkit/core'
import { AlignmentType, Paragraph, TextRun } from 'docx'

/**
 * Options for the Watermark plugin.
 */
export interface WatermarkOptions {
  /**
   * Watermark text.
   */
  text: string
  /**
   * Horizontal alignment.
   * @default 'center'
   */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  /**
   * Text color in hex RRGGBB.
   * @default 'BFBFBF'
   */
  color?: string
  /**
   * Font size in half-points (20 = 10pt).
   * @default 48
   */
  fontSize?: number
}

/**
 * Create a Watermark plugin instance.
 *
 * @returns A configured DocxPlugin for `'watermark'`
 *
 * @example
 * ```ts
 * import { createDocx, watermarkPlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(watermarkPlugin())
 *   .plugin('watermark', { text: 'DRAFT', color: 'FF0000' })
 * ```
 */
export function watermarkPlugin() {
  return definePlugin<'watermark', WatermarkOptions>({
    name: 'watermark',
    render(options) {
      return new Paragraph({
        alignment: options.alignment ?? AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
        children: [
          new TextRun({
            bold: true,
            color: options.color ?? 'BFBFBF',
            font: 'Arial',
            size: options.fontSize ?? 48,
            text: options.text,
          }),
        ],
      })
    },
  })
}
