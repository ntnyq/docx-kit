/**
 * Page Number plugin — generates a `Paragraph` containing a page number
 * field for placement inside section headers or footers.
 *
 * The paragraph can be inserted into a `Header` or `Footer` definition
 * and will render the current page number.
 *
 * @module plugins/page-number
 *
 * @example
 * ```ts
 * const pn = doc.plugin('pageNumber', { format: 'roman' })
 * doc.section({ footers: { default: [pn] } })
 * ```
 */

import { AlignmentType, PageNumber, Paragraph, TextRun } from 'docx'
import { definePlugin } from '../../types/plugin'

/** Options for the Page Number plugin. */
export interface PageNumberOptions {
  /**
   * Horizontal alignment of the page number.
   * @default 'center'
   */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  /**
   * Font size in half-points.
   * @default 20
   */
  fontSize?: number
  /**
   * Show "Page X of Y" instead of just "X".
   * @default false
   */
  showTotal?: boolean
}

/**
 * Create a Page Number plugin instance.
 *
 * @returns A configured DocxPlugin for `'pageNumber'`
 *
 * @example
 * ```ts
 * import { createDocx, pageNumberPlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(pageNumberPlugin())
 *   // Use the page number content in a footer
 *   const pn = doc.plugin('pageNumber', { alignment: 'center' })
 * ```
 */
export function pageNumberPlugin() {
  return definePlugin<'pageNumber', PageNumberOptions>({
    name: 'pageNumber',
    render(options) {
      const alignment = options.alignment ?? AlignmentType.CENTER
      const fontSize = options.fontSize ?? 20

      if (options.showTotal) {
        return new Paragraph({
          alignment,
          spacing: { after: 0, before: 0 },
          children: [
            new TextRun({
              font: 'Arial',
              size: fontSize,
              text: 'Page ',
            }),
            new TextRun({
              children: [PageNumber.CURRENT],
              font: 'Arial',
              size: fontSize,
            }),
            new TextRun({
              font: 'Arial',
              size: fontSize,
              text: ' of ',
            }),
            new TextRun({
              children: [PageNumber.TOTAL_PAGES],
              font: 'Arial',
              size: fontSize,
            }),
          ],
        })
      }

      return new Paragraph({
        alignment,
        spacing: { after: 0, before: 0 },
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            font: 'Arial',
            size: fontSize,
          }),
        ],
      })
    },
  })
}
