/**
 * Meeting Minutes plugin — structured meeting notes with header, meta info,
 * and an agenda table.
 *
 * Renders a title paragraph, a date/attendees summary line, and a 4-column
 * agenda table (Topic | Discussion | Decision | Owner).
 *
 * @module plugins/meeting-minutes
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(meetingMinutesPlugin())
 *   .plugin('meetingMinutes', {
 *     title: '项目周会纪要',
 *     date: '2026-06-11',
 *     attendees: ['张三', '李四', '王五'],
 *     agenda: [
 *       { topic: '项目进度', discussion: '模块A已完成80%', decision: '下周一上线', owner: '张三' },
 *       { topic: '风险项', discussion: '第三方API不稳定', decision: '增加重试机制', owner: '李四' },
 *     ],
 *   })
 *   .save('minutes.docx')
 * ```
 */

import {
  AlignmentType,
  BorderStyle,
  HeadingLevel,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import { definePlugin } from '../../types/plugin'
import type { IShadingAttributesProperties } from 'docx'

const HEADER_SHADING: IShadingAttributesProperties = {
  fill: '4472C4',
  type: ShadingType.CLEAR,
}

const BORDER = {
  bottom: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
  left: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
  right: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
  top: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
}

const AGENDA_COLUMNS = ['议题', '讨论', '决议', '负责人']

/** A single agenda item. */
export interface AgendaItem {
  /** Discussion notes. */
  discussion: string
  /** Meeting topic. */
  topic: string
  /** Decision made. */
  decision?: string
  /** Responsible person. */
  owner?: string
}

/** Options for the MeetingMinutes plugin. */
export interface MeetingMinutesOptions {
  /** Agenda items to render in the table. */
  agenda: AgendaItem[]
  /** Attendee names. */
  attendees: string[]
  /** Meeting date (e.g. "2026-06-11"). */
  date: string
  /** Meeting title (rendered as Heading 1). */
  title: string
}

/**
 * Create a MeetingMinutes plugin instance.
 *
 * @returns A configured DocxPlugin for `'meetingMinutes'`
 */
export function meetingMinutesPlugin() {
  return definePlugin<'meetingMinutes', MeetingMinutesOptions>({
    name: 'meetingMinutes',
    render(options) {
      // Title
      const titlePara = new Paragraph({
        children: [new TextRun({ bold: true, size: 32, text: options.title })],
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 200 },
      })

      // Meta info
      const metaPara = new Paragraph({
        spacing: { after: 300 },
        children: [
          new TextRun({ bold: true, text: '日期：' }),
          new TextRun({ text: options.date }),
          new TextRun({ text: '    ' }),
          new TextRun({ bold: true, text: '参会人员：' }),
          new TextRun({ text: options.attendees.join('、') }),
        ],
      })

      // Only add table if there are agenda items
      if (options.agenda.length === 0) {
        return [titlePara, metaPara]
      }

      // Header row
      const headerCells = AGENDA_COLUMNS.map(
        col =>
          new TableCell({
            borders: BORDER,
            shading: HEADER_SHADING,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ bold: true, color: 'FFFFFF', text: col }),
                ],
              }),
            ],
          }),
      )

      // Data rows
      const agendaRows = options.agenda.map(item => {
        const values = [
          item.topic,
          item.discussion,
          item.decision ?? '',
          item.owner ?? '',
        ]
        const cells = values.map(
          val =>
            new TableCell({
              borders: BORDER,
              children: [
                new Paragraph({
                  children: [new TextRun({ text: val })],
                }),
              ],
            }),
        )
        return new TableRow({ children: cells })
      })

      const rows = [new TableRow({ children: headerCells }), ...agendaRows]

      const agendaTable = new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      })

      return [titlePara, metaPara, agendaTable]
    },
  })
}
