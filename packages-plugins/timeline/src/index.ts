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

import { definePlugin } from '@docxkit/core'
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

/** A single timeline event. */
export interface TimelineEvent {
  /** Date / time label (e.g. "2026-06" or "Q3"). */
  date: string
  /** Short event title. */
  title: string
  /** Optional longer description. */
  description?: string
}

/** Options for the Timeline plugin. */
export interface TimelineOptions {
  /** Array of timeline events in chronological order. */
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
}

/** Default table column widths (in DXA, twips). */
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
export function timelinePlugin() {
  return definePlugin<'timeline', TimelineOptions>({
    name: 'timeline',
    render(options) {
      const { accentColor = '4472C4', events, layout = 'alternating' } = options

      if (events.length === 0) {
        return new Paragraph({ text: '(No events)' })
      }

      // Create an empty cell with no border (for layout spacing)
      const emptyCell = (width: number) =>
        new TableCell({
          children: [new Paragraph({ children: [] })],
          width: { size: width, type: WidthType.DXA },
          borders: {
            bottom: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          },
        })

      const connectorCell = () =>
        new TableCell({
          width: { size: COL_WIDTHS.connector, type: WidthType.DXA },
          borders: {
            bottom: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  color: accentColor,
                  font: 'Arial',
                  size: 24,
                  text: '\u25CF',
                }),
              ],
            }),
          ],
        })

      const dateCell = (date: string) =>
        new TableCell({
          width: { size: COL_WIDTHS.date, type: WidthType.DXA },
          borders: {
            bottom: { color: accentColor, size: 1, style: BorderStyle.SINGLE },
            left: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            right: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
            top: { color: 'FFFFFF', size: 0, style: BorderStyle.NONE },
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  bold: true,
                  color: accentColor,
                  font: 'Arial',
                  size: 22,
                  text: date,
                }),
              ],
            }),
          ],
          shading: {
            fill: 'F0F5FF',
            type: ShadingType.CLEAR,
          },
        })

      const contentCell = (title: string, description?: string) => {
        const runs: TextRun[] = [
          new TextRun({
            bold: true,
            font: 'Arial',
            size: 24,
            text: title,
          }),
        ]

        if (description) {
          runs.push(
            new TextRun({
              break: 1,
              color: '555555',
              font: 'Arial',
              size: 20,
              text: description,
            }),
          )
        }

        return new TableCell({
          children: [new Paragraph({ children: runs })],
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
        })
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
            children: [
              dateCell(event.date),
              connectorCell(),
              contentCell(event.title, event.description),
            ],
          })
        }

        return new TableRow({
          children: [
            emptyCell(COL_WIDTHS.date),
            connectorCell(),
            contentCell(event.title, event.description),
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
      })
    },
  })
}
