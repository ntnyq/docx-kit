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

import type {
  CellStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

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
 * Defaults shared by all calls using this plugin instance.
 */
export interface MeetingMinutesConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: MeetingMinutesStyles
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
  /**
   * Per-part styles for this invocation.
   */
  styles?: MeetingMinutesStyles
}

/**
 * Independently configurable MeetingMinutes parts.
 */
export interface MeetingMinutesStyles {
  /**
   * Body cell shading, borders and padding.
   */
  bodyCell?: Omit<CellStyleRule, 'height'>
  /**
   * Body paragraph alignment, spacing and flow.
   */
  bodyParagraph?: ParagraphStyleRule
  /**
   * Body text run typography.
   */
  bodyText?: TextStyleRule
  /**
   * Header cell shading, borders and padding.
   */
  headerCell?: Omit<CellStyleRule, 'height'>
  /**
   * Header paragraph alignment, spacing and flow.
   */
  headerParagraph?: ParagraphStyleRule
  /**
   * Header text run typography.
   */
  headerText?: TextStyleRule
  /**
   * Meta label run typography.
   */
  metaLabel?: TextStyleRule
  /**
   * Meta paragraph alignment, spacing and flow.
   */
  metaParagraph?: ParagraphStyleRule
  /**
   * Meta text run typography.
   */
  metaText?: TextStyleRule
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
 * Create a MeetingMinutes plugin instance.
 *
 * @returns A configured DocxPlugin for `'meetingMinutes'`
 */
export function meetingMinutesPlugin(pluginConfig: MeetingMinutesConfig = {}) {
  return definePlugin<'meetingMinutes', MeetingMinutesOptions>({
    name: 'meetingMinutes',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof MeetingMinutesStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const bodyCellStyle = partStyle('bodyCell')
      const bodyParagraphStyle = partStyle('bodyParagraph')
      const bodyTextStyle = partStyle('bodyText')
      const headerCellStyle = partStyle('headerCell')
      const headerParagraphStyle = partStyle('headerParagraph')
      const headerTextStyle = partStyle('headerText')
      const metaLabelStyle = partStyle('metaLabel')
      const metaParagraphStyle = partStyle('metaParagraph')
      const metaTextStyle = partStyle('metaText')
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
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
        titleParagraphStyle.paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { after: 200 },
          children: [
            new TextRun(
              titleStyle.text({ bold: true, size: 32, text: options.title }),
            ),
          ],
        }),
      )

      // Meta info
      const metaPara = new Paragraph(
        metaParagraphStyle.paragraph({
          spacing: { after: 300 },
          children: [
            new TextRun(metaLabelStyle.text({ bold: true, text: '日期：' })),
            new TextRun(metaTextStyle.text({ text: options.date })),
            new TextRun(metaTextStyle.text({ text: ' '.repeat(4) })),
            new TextRun(
              metaLabelStyle.text({ bold: true, text: '参会人员：' }),
            ),
            new TextRun(
              metaTextStyle.text({ text: options.attendees.join('、') }),
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
            headerCellStyle.cell({
              borders,
              shading: { fill: surface, type: ShadingType.CLEAR },
              children: [
                new Paragraph(
                  headerParagraphStyle.paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { after: 0, before: 0, line: 288 },
                    children: [
                      new TextRun(
                        headerTextStyle.text({
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
              bodyCellStyle.cell({
                borders,
                children: [
                  new Paragraph(
                    bodyParagraphStyle.paragraph({
                      children: [
                        new TextRun(bodyTextStyle.text({ text: val })),
                      ],
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
