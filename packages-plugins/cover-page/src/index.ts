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
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  TextRun,
} from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface CoverPageConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: CoverPageStyles
}

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
   * Per-part styles for this invocation.
   */
  styles?: CoverPageStyles
  /**
   * Sub-title text displayed below the title.
   */
  subtitle?: string
}

/**
 * Independently configurable CoverPage parts.
 */
export interface CoverPageStyles {
  /**
   * Author run typography.
   */
  author?: TextStyleRule
  /**
   * Author paragraph alignment, spacing and flow.
   */
  authorParagraph?: ParagraphStyleRule
  /**
   * Date run typography.
   */
  date?: TextStyleRule
  /**
   * Date paragraph alignment, spacing and flow.
   */
  dateParagraph?: ParagraphStyleRule
  /**
   * Organization run typography.
   */
  organization?: TextStyleRule
  /**
   * Organization paragraph alignment, spacing and flow.
   */
  organizationParagraph?: ParagraphStyleRule
  /**
   * Rule paragraph alignment, spacing and flow.
   */
  ruleParagraph?: ParagraphStyleRule
  /**
   * Subtitle run typography.
   */
  subtitle?: TextStyleRule
  /**
   * Subtitle paragraph alignment, spacing and flow.
   */
  subtitleParagraph?: ParagraphStyleRule
  /**
   * Title run typography.
   */
  title?: TextStyleRule
  /**
   * Title paragraph alignment, spacing and flow.
   */
  titleParagraph?: ParagraphStyleRule
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
export function coverPagePlugin(pluginConfig: CoverPageConfig = {}) {
  return definePlugin<'coverPage', CoverPageOptions>({
    name: 'coverPage',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof CoverPageStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const authorStyle = partStyle('author')
      const authorParagraphStyle = partStyle('authorParagraph')
      const dateStyle = partStyle('date')
      const dateParagraphStyle = partStyle('dateParagraph')
      const organizationStyle = partStyle('organization')
      const organizationParagraphStyle = partStyle('organizationParagraph')
      const ruleParagraphStyle = partStyle('ruleParagraph')
      const subtitleStyle = partStyle('subtitle')
      const subtitleParagraphStyle = partStyle('subtitleParagraph')
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
      const colors = config.theme?.colors
      const alignment = options.alignment ?? AlignmentType.CENTER
      const showRule = options.showRule === true

      const titlePara = new Paragraph(
        titleParagraphStyle.paragraph({
          alignment,
          keepNext: true,
          spacing: { after: 240, before: 1800 },
          style: 'Title',
          children: [
            new TextRun(
              titleStyle.text({
                bold: true,
                size: 64, // 32pt
                text: options.title,
              }),
            ),
          ],
          ...(options.backgroundColor
            ? {
                shading: {
                  fill: normalizeHexColor(options.backgroundColor),
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
                subtitleParagraphStyle.paragraph({
                  alignment,
                  spacing: { after: 400 },
                  children: [
                    new TextRun(
                      subtitleStyle.text({
                        color: normalizeHexColor(colors?.muted) ?? '666666',
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
                ruleParagraphStyle.paragraph({
                  alignment,
                  children: [],
                  spacing: { after: 400 },
                  border: {
                    bottom: {
                      color: normalizeHexColor(colors?.accent) ?? '2E75B6',
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
                authorParagraphStyle.paragraph({
                  alignment,
                  spacing: { after: 120 },
                  children: [
                    new TextRun(
                      authorStyle.text({
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
                organizationParagraphStyle.paragraph({
                  alignment,
                  spacing: { after: 120 },
                  children: [
                    new TextRun(
                      organizationStyle.text({
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
                dateParagraphStyle.paragraph({
                  alignment,
                  spacing: { after: 0 },
                  children: [
                    new TextRun(
                      dateStyle.text({
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
