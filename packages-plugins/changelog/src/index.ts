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
    render(options) {
      const title = new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            bold: true,
            size: 28,
            text: options.title ?? 'Changelog',
          }),
        ],
      })

      if (options.entries.length === 0) {
        return [
          title,
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ color: '999999', text: '(no entries)' })],
          }),
        ]
      }

      const header = new TableRow({
        children: ['Version', 'Date', 'Type', 'Changes'].map(
          label =>
            new TableCell({
              shading: { fill: '4472C4', type: ShadingType.CLEAR },
              borders: {
                bottom: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
                left: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
                right: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
                top: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ bold: true, color: 'FFFFFF', text: label }),
                  ],
                }),
              ],
            }),
        ),
      })

      const rows = options.entries.map(
        entry =>
          new TableRow({
            children: [
              entry.version,
              entry.date,
              entry.type,
              entry.changes,
            ].map(
              (value, index) =>
                new TableCell({
                  borders: {
                    bottom: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    left: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    right: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    top: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                  },
                  children: [
                    new Paragraph({ children: [new TextRun({ text: value })] }),
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
          }),
      )

      return [
        title,
        new Table({
          rows: [header, ...rows],
          width: { size: 100, type: WidthType.PERCENTAGE },
        }),
      ]
    },
  })
}
