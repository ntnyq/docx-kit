import { definePlugin } from '@docxkit/core'
import { HeadingLevel, Paragraph, TableOfContents, TextRun } from 'docx'

export interface TocOptions {
  maxLevel?: number
  title?: string
}

export function tocPlugin() {
  return definePlugin<'toc', TocOptions>({
    name: 'toc',
    render(options) {
      const maxLevel = options.maxLevel ?? 3
      const range = `1-${maxLevel}`

      return [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
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
