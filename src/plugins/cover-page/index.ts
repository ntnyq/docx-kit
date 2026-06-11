/**
 * Cover Page plugin — generates a professional title page.
 *
 * Renders a centered block with title, subtitle, author, date,
 * organization, and an optional decorative horizontal rule.
 *
 * @module plugins/cover-page
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(coverPagePlugin())
 *   .plugin('coverPage', {
 *     title: 'Q3 运营报告',
 *     subtitle: '数据驱动 · 智能决策',
 *     author: '数据分析部',
 *   })
 *   .save('report.docx')
 * ```
 */

import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  TextRun,
} from 'docx'
import { definePlugin } from '../../types/plugin'

/** Options for the Cover Page plugin. */
export interface CoverPageOptions {
  /** Main title text (required). */
  title: string
  /** Horizontal text alignment. @default CENTER */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  /** Author or department name. */
  author?: string
  /** Background color of the cover page in hex RRGGBB. */
  backgroundColor?: string
  /** Date string (e.g. "2026-06-11"). */
  date?: string
  /** Organization / company name. */
  organization?: string
  /**
   * Show a decorative horizontal rule between title and author.
   * @default true
   */
  showRule?: boolean
  /** Sub-title text displayed below the title. */
  subtitle?: string
}

/**
 * Create a Cover Page plugin instance.
 *
 * @returns A configured DocxPlugin for `'coverPage'`
 *
 * @example
 * ```ts
 * import { coverPagePlugin, createDocx } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(coverPagePlugin())
 *   .plugin('coverPage', {
 *     title: '年度报告',
 *     subtitle: '2026',
 *     author: '战略发展部',
 *     organization: 'XX 科技集团',
 *   })
 * ```
 */
export function coverPagePlugin() {
  return definePlugin<'coverPage', CoverPageOptions>({
    name: 'coverPage',
    render(options) {
      const alignment = options.alignment ?? AlignmentType.CENTER
      const showRule = options.showRule !== false

      const titlePara = new Paragraph({
        alignment,
        spacing: { after: 200 },
        children: [
          new TextRun({
            bold: true,
            font: 'Arial',
            size: 56, // 28pt
            text: options.title,
          }),
        ],
        ...(options.backgroundColor
          ? {
              shading: {
                fill: options.backgroundColor,
                type: ShadingType.CLEAR,
              },
            }
          : {}),
      })

      return [
        new Paragraph({ children: [], spacing: { before: 2400 } }),
        titlePara,
        ...(options.subtitle
          ? [
              new Paragraph({
                alignment,
                spacing: { after: 400 },
                children: [
                  new TextRun({
                    color: '555555',
                    font: 'Arial',
                    size: 32,
                    text: options.subtitle,
                  }),
                ],
              }),
            ]
          : []),
        ...(showRule
          ? [
              new Paragraph({
                alignment,
                children: [],
                spacing: { after: 400 },
                border: {
                  bottom: {
                    color: '4472C4',
                    size: 6,
                    space: 1,
                    style: BorderStyle.SINGLE,
                  },
                },
              }),
            ]
          : []),
        ...(options.author
          ? [
              new Paragraph({
                alignment,
                spacing: { after: 120 },
                children: [
                  new TextRun({
                    font: 'Arial',
                    size: 28,
                    text: options.author,
                  }),
                ],
              }),
            ]
          : []),
        ...(options.organization
          ? [
              new Paragraph({
                alignment,
                spacing: { after: 120 },
                children: [
                  new TextRun({
                    color: '333333',
                    font: 'Arial',
                    size: 24,
                    text: options.organization,
                  }),
                ],
              }),
            ]
          : []),
        ...(options.date
          ? [
              new Paragraph({
                alignment,
                spacing: { after: 0 },
                children: [
                  new TextRun({
                    color: '777777',
                    font: 'Arial',
                    size: 24,
                    text: options.date,
                  }),
                ],
              }),
            ]
          : []),
      ]
    },
  })
}
