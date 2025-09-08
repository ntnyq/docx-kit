import { Document, Footer, Header, Packer, Paragraph, TextRun } from 'docx'
import type { ISectionOptions, Table } from 'docx'
import type { DocumentOptions, TitleOptions } from './types'

export class DocxBuilder {
  #children: Array<Paragraph | Table> = []
  #footer: Footer | undefined
  #header: Header | undefined
  #options: DocumentOptions

  addFooter(text: string): this {
    this.#footer = new Footer({
      children: [
        new Paragraph({
          text,
        }),
      ],
    })
    return this
  }

  addHeader(text: string): this {
    this.#header = new Header({
      children: [
        new Paragraph({
          text,
        }),
      ],
    })
    return this
  }

  addPageBreak(): this {
    this.#children.push(
      new Paragraph({
        children: [],
        pageBreakBefore: true,
      }),
    )
    return this
  }

  addParagraph(text: string) {
    const paragraph = new Paragraph({
      children: [
        new TextRun({
          text,
        }),
      ],
    })
    this.#children.push(paragraph)
    return this
  }

  addRaw(content: Paragraph | Table): this {
    this.#children.push(content)
    return this
  }

  addTitle(text: string, options: TitleOptions = {}) {
    const level = options.level ?? 1
    const heading = new Paragraph({
      alignment: options.alignment,
      heading: `HEADING_${level}` as any,
      spacing: options.spacing,
      text,
      children: [
        new TextRun({
          bold: true,
          color: options.color,
          font: options.font,
          size: options.fontSize ? options.fontSize * 2 : undefined,
          text,
        }),
      ],
    })
    this.#children.push(heading)
    return this
  }

  private buildDocument() {
    const section: ISectionOptions = {
      children: this.#children,
      footers: this.#footer ? { default: this.#footer } : undefined,
      headers: this.#header ? { default: this.#header } : undefined,
    }

    return new Document({
      creator: this.#options.creator,
      description: this.#options.description,
      sections: [section],
      title: this.#options.title,
    })
  }

  constructor(options: DocumentOptions = {}) {
    this.#options = options
  }

  async getBlob(): Promise<Blob> {
    const doc = this.buildDocument()
    return await Packer.toBlob(doc)
  }

  toDocx(): Document {
    return this.buildDocument()
  }
}
