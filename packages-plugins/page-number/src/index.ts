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

import { definePlugin } from '@docxkit/core'
import { AlignmentType, PageNumber, Paragraph, TextRun } from 'docx'

/** Options for the Page Number plugin. */
export interface PageNumberOptions {
  /**
   * Horizontal alignment of the page number.
   * @default 'center'
   */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  /**
   * Font size in **points** (will be converted to docx half-points internally).
   * @default 10
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
      // docx's `TextRun.size` is in half-points; user-facing option is in points.
      const halfPoints = (options.fontSize ?? 10) * 2

      // Per the docx library docs, `PageNumber.CURRENT` and `PageNumber.TOTAL_PAGES`
      // must be placed inside a single `TextRun.children` array — splitting them across
      // multiple `TextRun`s breaks the field rendering in Word.
      const children: (string | typeof PageNumber.CURRENT)[] = options.showTotal
        ? ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES]
        : [PageNumber.CURRENT]

      return new Paragraph({
        alignment,
        spacing: { after: 0, before: 0 },
        children: [
          new TextRun({
            children,
            font: 'Arial',
            size: halfPoints,
          }),
        ],
      })
    },
  })
}
