/**
 * Business report template.
 *
 * Generates a professional report with cover page,
 * executive summary, body sections, and conclusion.
 *
 * @module ai/templates/report
 */

import type { BlockNode, DocxSchema } from '@docxkit/core'
import type { AiTemplate, AiTemplateSchema } from '../types'

/**
 * Report template parameters.
 */
export interface ReportParams {
  /**
   * Report title.
   */
  title: string
  /**
   * Author name(s).
   */
  author?: string
  /**
   * Conclusion text.
   */
  conclusion?: string
  /**
   * Report date (e.g. "2026-06-12").
   */
  date?: string
  /**
   * Executive summary paragraphs.
   */
  executiveSummary?: string[]
  /**
   * Full body sections.
   */
  sections?: {
    /**
     * Section body paragraphs.
     */
    content?: string[]
    /**
     * Section heading.
     */
    heading?: string
  }[]
}

const systemPrompt = `You are a professional document generator. Generate a docx-kit JSON schema for a business report document.

The report should include:
1. A cover page with title, author, and date
2. An executive summary section
3. Multiple body sections with headings and content
4. A conclusion section

Use the following docx-kit node types:
- { type: "heading", level: N, text: "..." } for section headings
- { type: "paragraph", text: "..." } for body text
- { type: "pageBreak" } for page breaks between sections
- { type: "heading", level: 1, text: "..." } for the cover title

Keep the content professional, concise, and well-structured.
`

const schema: AiTemplateSchema = {
  title: 'ReportParams',
  type: 'object',
  description:
    'Business report with cover page, executive summary, body sections, and conclusion',
  properties: {
    author: { description: 'Author name(s)', type: 'string' },
    conclusion: { description: 'Conclusion text', type: 'string' },
    date: { description: 'Report date (e.g. "2026-06-12")', type: 'string' },
    title: { description: 'Report title', required: true, type: 'string' },
    executiveSummary: {
      description: 'Executive summary paragraphs',
      items: { type: 'string' },
      type: 'array',
    },
    sections: {
      description: 'Body sections',
      type: 'array',
      items: {
        type: 'object',
        properties: {
          heading: { description: 'Heading', type: 'string' },
          content: {
            description: 'Paragraphs',
            items: { type: 'string' },
            type: 'array',
          },
        },
      },
    },
  },
}

/**
 * Generate a business report DocxSchema from template parameters.
 *
 * Produces a document with a cover page, executive summary,
 * body sections with headings, and a conclusion.
 *
 * @param params - — Report template parameters
 * @returns A DocxSchema ready for rendering
 */
function generate(params: ReportParams): DocxSchema {
  const content: BlockNode[] = [
    { level: 1, text: params.title, type: 'heading' },
  ]
  if (params.author) {
    content.push({ text: params.author, type: 'paragraph' })
  }
  if (params.date) {
    content.push({ text: params.date, type: 'paragraph' })
  }
  content.push({ type: 'pageBreak' })

  if (params.executiveSummary && params.executiveSummary.length > 0) {
    content.push({ level: 1, text: 'Executive Summary', type: 'heading' })
    for (const text of params.executiveSummary) {
      content.push({ text, type: 'paragraph' })
    }
  }

  if (params.sections) {
    for (const section of params.sections) {
      if (section.heading) {
        content.push({ level: 2, text: section.heading, type: 'heading' })
      }
      if (section.content) {
        for (const text of section.content) {
          content.push({ text, type: 'paragraph' })
        }
      }
    }
  }

  if (params.conclusion) {
    content.push(
      { level: 2, text: 'Conclusion', type: 'heading' },
      { text: params.conclusion, type: 'paragraph' },
    )
  }

  return { content }
}

export const reportTemplate: AiTemplate<ReportParams> = {
  generate,
  name: 'report',
  schema,
  systemPrompt,
  description:
    'Business report with cover page, executive summary, body sections, and conclusion',
}
