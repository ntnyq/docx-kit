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
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
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

const AGENDA_COLUMNS = ['议题', '讨论', '决议', '负责人']

/**
 * A single agenda item.
 */
export interface AgendaItem {
  /**
   * Discussion notes.
   */
  discussion: string
  /**
   * Meeting topic.
   */
  topic: string
  /**
   * Decision made.
   */
  decision?: string
  /**
   * Responsible person.
   */
  owner?: string
}

/**
 * Options for the MeetingMinutes plugin.
 */
export interface MeetingMinutesOptions {
  /**
   * Agenda items to render in the table.
   */
  agenda: AgendaItem[]
  /**
   * Attendee names.
   */
  attendees: string[]
  /**
   * Meeting date (e.g. "2026-06-11").
   */
  date: string
  /**
   * Meeting title (rendered as Heading 1).
   */
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
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
      const colors = config.theme?.colors
      const border = {
        color: normalizeHexColor(colors?.border) ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const borders = {
        bottom: border,
        left: border,
        right: border,
        top: border,
      }
      const surface = normalizeHexColor(colors?.surface) ?? 'F4F6F8'
      const primary = normalizeHexColor(colors?.primary) ?? '1B2A4A'
      // Title
      const titlePara = new Paragraph(
        pluginStyles.paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { after: 200 },
          children: [
            new TextRun(
              pluginStyles.text({ bold: true, size: 32, text: options.title }),
            ),
          ],
        }),
      )

      // Meta info
      const metaPara = new Paragraph(
        pluginStyles.paragraph({
          spacing: { after: 300 },
          children: [
            new TextRun(pluginStyles.text({ bold: true, text: '日期：' })),
            new TextRun(pluginStyles.text({ text: options.date })),
            new TextRun(pluginStyles.text({ text: ' '.repeat(4) })),
            new TextRun(pluginStyles.text({ bold: true, text: '参会人员：' })),
            new TextRun(
              pluginStyles.text({ text: options.attendees.join('、') }),
            ),
          ],
        }),
      )

      // Only add table if there are agenda items
      if (options.agenda.length === 0) {
        return [titlePara, metaPara]
      }

      // Header row
      const headerCells = AGENDA_COLUMNS.map(
        col =>
          new TableCell(
            pluginStyles.cell({
              borders,
              shading: { fill: surface, type: ShadingType.CLEAR },
              children: [
                new Paragraph(
                  pluginStyles.paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { after: 0, before: 0, line: 288 },
                    children: [
                      new TextRun(
                        pluginStyles.text({
                          bold: true,
                          color: primary,
                          text: col,
                        }),
                      ),
                    ],
                  }),
                ),
              ],
            }),
          ),
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
            new TableCell(
              pluginStyles.cell({
                borders,
                children: [
                  new Paragraph(
                    pluginStyles.paragraph({
                      children: [new TextRun(pluginStyles.text({ text: val }))],
                    }),
                  ),
                ],
              }),
            ),
        )
        return new TableRow({
          cantSplit: true,
          children: cells,
        })
      })

      const rows = [
        new TableRow({
          cantSplit: true,
          children: headerCells,
          tableHeader: true,
        }),
        ...agendaRows,
      ]

      const agendaTable = new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: {
          bottom: 100,
          left: 140,
          right: 140,
          top: 100,
        },
      })

      return [titlePara, metaPara, agendaTable]
    },
  })
}
