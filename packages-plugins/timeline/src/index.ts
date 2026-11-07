/**
 * Timeline plugin — renders a chronological timeline as a styled table.
 *
 * Each event occupies one table row: a date cell, a visual connector
 * cell, and a content cell (title + optional description).
 *
 * @module plugins/timeline
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(timelinePlugin())
 *   .plugin('timeline', {
 *     events: [
 *       { date: '2026-01', title: '项目立项', description: '完成需求评审' },
 *       { date: '2026-03', title: 'MVP 发布', description: '核心功能上线' },
 *     ],
 *   })
 *   .save('timeline.docx')
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
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'

import type {
  CellStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface TimelineConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: TimelineStyles
}

/**
 * A single timeline event.
 */
export interface TimelineEvent {
  /**
   * Date / time label (e.g. "2026-06" or "Q3").
   */
  date: string
  /**
   * Short event title.
   */
  title: string
  /**
   * Optional longer description.
   */
  description?: string
}

/**
 * Options for the Timeline plugin.
 */
export interface TimelineOptions {
  /**
   * Array of timeline events in chronological order.
   */
  events: TimelineEvent[]
  /**
   * Accent color for the connector line and date highlight.
   * @default '4472C4'
   */
  accentColor?: string
  /**
   * Layout style.
   *
   * - `'alternating'` — dates alternate left / right (default)
   * - `'left'` — all dates on the left, content on the right
   * - `'right'` — all content on the left, dates on the right
   *
   * @default 'alternating'
   */
  layout?: 'alternating' | 'left' | 'right'
  /**
   * Per-part styles for this invocation.
   */
  styles?: TimelineStyles
}

/**
 * Independently configurable Timeline parts.
 */
export interface TimelineStyles {
  /**
   * Connector marker typography.
   */
  connector?: TextStyleRule
  /**
   * Connector cell shading, borders and padding.
   */
  connectorCell?: Omit<CellStyleRule, 'height'>
  /**
   * Connector paragraph alignment, spacing and flow.
   */
  connectorParagraph?: ParagraphStyleRule
  /**
   * Content cell shading, borders and padding.
   */
  contentCell?: Omit<CellStyleRule, 'height'>
  /**
   * Content paragraph alignment, spacing and flow.
   */
  contentParagraph?: ParagraphStyleRule
  /**
   * Date run typography.
   */
  date?: TextStyleRule
  /**
   * Date cell shading, borders and padding.
   */
  dateCell?: Omit<CellStyleRule, 'height'>
  /**
   * Date paragraph alignment, spacing and flow.
   */
  dateParagraph?: ParagraphStyleRule
  /**
   * Description run typography.
   */
  description?: TextStyleRule
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Title run typography.
   */
  title?: TextStyleRule
}

/**
 * Default table column widths (in DXA, twips).
 */
const COL_WIDTHS = {
  connector: 600,
  content: 6400,
  date: 1800,
} as const

/**
 * Create a Timeline plugin instance.
 *
 * @returns A configured DocxPlugin for `'timeline'`
 *
 * @example
 * ```ts
 * import { createDocx, timelinePlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(timelinePlugin())
 *   .plugin('timeline', {
 *     events: [
 *       { date: '2026-01', title: '项目启动', description: '团队组建完成' },
 *       { date: '2026-04', title: 'Beta 测试', description: '收集用户反馈' },
 *       { date: '2026-06', title: '正式发布' },
 *     ],
 *   })
 * ```
 */
export function timelinePlugin(pluginConfig: TimelineConfig = {}) {
  return definePlugin<'timeline', TimelineOptions>({
    name: 'timeline',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof TimelineStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const connectorStyle = partStyle('connector')
      const connectorCellStyle = partStyle('connectorCell')
      const connectorParagraphStyle = partStyle('connectorParagraph')
      const contentCellStyle = partStyle('contentCell')
      const contentParagraphStyle = partStyle('contentParagraph')
      const dateStyle = partStyle('date')
      const dateCellStyle = partStyle('dateCell')
      const dateParagraphStyle = partStyle('dateParagraph')
      const descriptionStyle = partStyle('description')
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const titleStyle = partStyle('title')
      const colors = config.theme?.colors
      const {
        accentColor = normalizeHexColor(colors?.accent) ?? '2E75B6',
        events,
        layout = 'alternating',
      } = options

      if (events.length === 0) {
        return new Paragraph(
          emptyParagraphStyle.paragraph({
            children: [
              new TextRun(emptyTextStyle.text({ text: '(No events)' })),
            ],
          }),
        )
      }

      const connectorCell = () =>
        new TableCell(
          connectorCellStyle.cell({
            width: { size: COL_WIDTHS.connector, type: WidthType.DXA },
            borders: {
              bottom: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            },
            children: [
              new Paragraph(
                connectorParagraphStyle.paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun(
                      connectorStyle.text({
                        color: normalizeHexColor(accentColor),
                        size: 24,
                        text: '\u{25CF}',
                      }),
                    ),
                  ],
                }),
              ),
            ],
          }),
        )

      const dateCell = (date: string) =>
        new TableCell(
          dateCellStyle.cell({
            width: { size: COL_WIDTHS.date, type: WidthType.DXA },
            borders: {
              left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              bottom: {
                color: normalizeHexColor(accentColor),
                size: 1,
                style: BorderStyle.SINGLE,
              },
            },
            children: [
              new Paragraph(
                dateParagraphStyle.paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun(
                      dateStyle.text({
                        bold: true,
                        color: normalizeHexColor(accentColor),
                        size: 22,
                        text: date,
                      }),
                    ),
                  ],
                }),
              ),
            ],
            shading: {
              fill: normalizeHexColor(colors?.surface) ?? 'F4F6F8',
              type: ShadingType.CLEAR,
            },
          }),
        )

      const contentCell = (title: string, description?: string) => {
        const runs: TextRun[] = [
          new TextRun(
            titleStyle.text({
              bold: true,
              size: 24,
              text: title,
            }),
          ),
        ]

        if (description) {
          runs.push(
            new TextRun(
              descriptionStyle.text({
                break: 1,
                color: '555555',
                size: 20,
                text: description,
              }),
            ),
          )
        }

        return new TableCell(
          contentCellStyle.cell({
            width: { size: COL_WIDTHS.content, type: WidthType.DXA },
            borders: {
              left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
              bottom: {
                color: 'DDDDDD',
                size: 1,
                style: BorderStyle.SINGLE,
              },
            },
            children: [
              new Paragraph(
                contentParagraphStyle.paragraph({ children: runs }),
              ),
            ],
          }),
        )
      }

      const rows: TableRow[] = events.map((event, index) => {
        const isLeft =
          layout === 'left'
            ? true
            : layout === 'right'
              ? false
              : index % 2 === 0

        if (isLeft) {
          return new TableRow({
            cantSplit: true,
            children: [
              dateCell(event.date),
              connectorCell(),
              contentCell(event.title, event.description),
            ],
          })
        }

        return new TableRow({
          cantSplit: true,
          children: [
            contentCell(event.title, event.description),
            connectorCell(),
            dateCell(event.date),
          ],
        })
      })

      return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          bottom: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          insideHorizontal: {
            color: 'FFFFFF',
            size: 0,
            style: BorderStyle.NONE,
          },
        },
        margins: {
          bottom: 100,
          left: 100,
          right: 100,
          top: 100,
        },
      })
    },
  })
}
