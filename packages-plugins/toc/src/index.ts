import { definePlugin } from '@docxkit/core'
import { Paragraph, TableOfContents, TextRun } from 'docx'

export interface TocOptions {
  maxLevel?: number
  title?: string
}

export function tocPlugin() {
  return definePlugin<'toc', TocOptions>({
    name: 'toc',
    render(options) {
      const maxLevel = Math.min(
        9,
        Math.max(1, Math.trunc(options.maxLevel ?? 3)),
      )
      const range = `1-${maxLevel}`

      return [
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({
              bold: true,
              size: 28,
              text: options.title ?? 'Contents',
            }),
          ],
        }),
        new TableOfContents(options.title ?? 'Contents', {
          beginDirty: true,
          headingStyleRange: range,
          hyperlink: true,
          pageNumbersEntryLevelsRange: range,
        }),
      ]
    },
  })
}
