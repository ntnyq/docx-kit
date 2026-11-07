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

import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { AlignmentType, Paragraph, TextRun } from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface WatermarkConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: WatermarkStyles
}

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
  /**
   * Per-part styles for this invocation.
   */
  styles?: WatermarkStyles
}

/**
 * Independently configurable Watermark parts.
 */
export interface WatermarkStyles {
  /**
   * Paragraph alignment, spacing and flow.
   */
  paragraph?: ParagraphStyleRule
  /**
   * Text run typography.
   */
  text?: TextStyleRule
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
export function watermarkPlugin(pluginConfig: WatermarkConfig = {}) {
  return definePlugin<'watermark', WatermarkOptions>({
    name: 'watermark',
    render(options, context) {
      const partStyle = (...parts: (keyof WatermarkStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const paragraphStyle = partStyle('paragraph')
      const textStyle = partStyle('text')
      return new Paragraph(
        paragraphStyle.paragraph({
          alignment: options.alignment ?? AlignmentType.CENTER,
          spacing: { after: 0, before: 0 },
          children: [
            new TextRun(
              textStyle.text({
                bold: true,
                color: normalizeHexColor(options.color ?? 'BFBFBF'),
                font: 'Arial',
                size: options.fontSize ?? 48,
                text: options.text,
              }),
            ),
          ],
        }),
      )
    },
  })
}
