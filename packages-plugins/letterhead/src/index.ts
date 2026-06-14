import { definePlugin } from '@docxkit/core'
import { AlignmentType, BorderStyle, Paragraph, TextRun } from 'docx'

export interface LetterheadOptions {
  companyName: string
  address?: string
  email?: string
  phone?: string
  tagline?: string
  website?: string
}

export function letterheadPlugin() {
  return definePlugin<'letterhead', LetterheadOptions>({
    name: 'letterhead',
    render(options) {
      const contact = [options.phone, options.email, options.website]
        .filter(Boolean)
        .join(' | ')

      return [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [
            new TextRun({ bold: true, size: 30, text: options.companyName }),
          ],
        }),
        ...(options.tagline
          ? [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [
                  new TextRun({ color: '666666', text: options.tagline }),
                ],
              }),
            ]
          : []),
        ...(contact
          ? [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: contact })],
                spacing: { after: 100 },
              }),
            ]
          : []),
        ...(options.address
          ? [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: options.address })],
                spacing: { after: 120 },
              }),
            ]
          : []),
        new Paragraph({
          spacing: { after: 220 },
          border: {
            bottom: {
              color: '4472C4',
              size: 6,
              space: 1,
              style: BorderStyle.SINGLE,
            },
          },
        }),
      ]
    },
  })
}
