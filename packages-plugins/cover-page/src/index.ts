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

import { createPluginStyle, definePlugin } from '@docxkit/core'
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  TextRun,
} from 'docx'

/**
 * Options for the Cover Page plugin.
 */
export interface CoverPageOptions {
  /**
   * Main title text (required).
   */
  title: string
  /**
   * Horizontal text alignment.
   * @default CENTER
   */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  /**
   * Author or department name.
   */
  author?: string
  /**
   * Background color of the cover page in hex RRGGBB.
   */
  backgroundColor?: string
  /**
   * Date string (e.g. "2026-06-11").
   */
  date?: string
  /**
   * Organization / company name.
   */
  organization?: string
  /**
   * Show a decorative horizontal rule between title and author.
   * @default false
   */
  showRule?: boolean
  /**
   * Sub-title text displayed below the title.
   */
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
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
      const colors = config.theme?.colors
      const alignment = options.alignment ?? AlignmentType.CENTER
      const showRule = options.showRule === true

      const titlePara = new Paragraph(
        pluginStyles.paragraph({
          alignment,
          keepNext: true,
          spacing: { after: 240, before: 1800 },
          style: 'Title',
          children: [
            new TextRun(
              pluginStyles.text({
                bold: true,
                size: 64, // 32pt
                text: options.title,
              }),
            ),
          ],
          ...(options.backgroundColor
            ? {
                shading: {
                  fill: options.backgroundColor.replace(/^#/, ''),
                  type: ShadingType.CLEAR,
                },
              }
            : {}),
        }),
      )

      return [
        titlePara,
        ...(options.subtitle
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  alignment,
                  spacing: { after: 400 },
                  children: [
                    new TextRun(
                      pluginStyles.text({
                        color: colors?.muted?.replace(/^#/, '') ?? '666666',
                        size: 28,
                        text: options.subtitle,
                      }),
                    ),
                  ],
                }),
              ),
            ]
          : []),
        ...(showRule
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  alignment,
                  children: [],
                  spacing: { after: 400 },
                  border: {
                    bottom: {
                      color: colors?.accent?.replace(/^#/, '') ?? '2E75B6',
                      size: 6,
                      space: 1,
                      style: BorderStyle.SINGLE,
                    },
                  },
                }),
              ),
            ]
          : []),
        ...(options.author
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  alignment,
                  spacing: { after: 120 },
                  children: [
                    new TextRun(
                      pluginStyles.text({
                        size: 28,
                        text: options.author,
                      }),
                    ),
                  ],
                }),
              ),
            ]
          : []),
        ...(options.organization
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  alignment,
                  spacing: { after: 120 },
                  children: [
                    new TextRun(
                      pluginStyles.text({
                        color: '333333',
                        size: 24,
                        text: options.organization,
                      }),
                    ),
                  ],
                }),
              ),
            ]
          : []),
        ...(options.date
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  alignment,
                  spacing: { after: 0 },
                  children: [
                    new TextRun(
                      pluginStyles.text({
                        color: '777777',
                        size: 24,
                        text: options.date,
                      }),
                    ),
                  ],
                }),
              ),
            ]
          : []),
      ]
    },
  })
}
