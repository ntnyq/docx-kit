import { createPluginStyle, definePlugin } from '@docxkit/core'
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

const TYPE_COLORS = {
  added: 'E7F6EA',
  changed: 'FFF4D6',
  fixed: 'E8F1FB',
  removed: 'FDE7E9',
} as const

export interface ChangelogEntry {
  changes: string
  date: string
  type: keyof typeof TYPE_COLORS
  version: string
}

export interface ChangelogOptions {
  entries: ChangelogEntry[]
  title?: string
}

export function changelogPlugin() {
  return definePlugin<'changelog', ChangelogOptions>({
    name: 'changelog',
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
      const colors = config.theme?.colors
      const border = {
        color: colors?.border?.replace(/^#/, '') ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const borders = {
        bottom: border,
        left: border,
        right: border,
        top: border,
      }
      const surface = colors?.surface?.replace(/^#/, '') ?? 'F4F6F8'
      const primary = colors?.primary?.replace(/^#/, '') ?? '1B2A4A'
      const title = new Paragraph(
        pluginStyles.paragraph({
          spacing: { after: 180 },
          children: [
            new TextRun(
              pluginStyles.text({
                bold: true,
                size: 28,
                text: options.title ?? 'Changelog',
              }),
            ),
          ],
        }),
      )

      if (options.entries.length === 0) {
        return [
          title,
          new Paragraph(
            pluginStyles.paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun(
                  pluginStyles.text({ color: '999999', text: '(no entries)' }),
                ),
              ],
            }),
          ),
        ]
      }

      const header = new TableRow({
        cantSplit: true,
        tableHeader: true,
        children: ['Version', 'Date', 'Type', 'Changes'].map(
          label =>
            new TableCell(
              pluginStyles.cell({
                borders,
                shading: { fill: surface, type: ShadingType.CLEAR },
                children: [
                  new Paragraph(
                    pluginStyles.paragraph({
                      children: [
                        new TextRun(
                          pluginStyles.text({
                            bold: true,
                            color: primary,
                            text: label,
                          }),
                        ),
                      ],
                    }),
                  ),
                ],
              }),
            ),
        ),
      })

      const rows = options.entries.map(
        entry =>
          new TableRow({
            cantSplit: true,
            children: [
              entry.version,
              entry.date,
              entry.type,
              entry.changes,
            ].map(
              (value, index) =>
                new TableCell(
                  pluginStyles.cell({
                    borders,
                    children: [
                      new Paragraph(
                        pluginStyles.paragraph({
                          children: [
                            new TextRun(pluginStyles.text({ text: value })),
                          ],
                        }),
                      ),
                    ],
                    shading:
                      index === 2
                        ? {
                            fill: TYPE_COLORS[entry.type],
                            type: ShadingType.CLEAR,
                          }
                        : undefined,
                  }),
                ),
            ),
          }),
      )

      return [
        title,
        new Table({
          margins: { bottom: 100, left: 140, right: 140, top: 100 },
          rows: [header, ...rows],
          width: { size: 100, type: WidthType.PERCENTAGE },
        }),
      ]
    },
  })
}
