/**
 * Formal letter template.
 *
 * Generates a professional letter with sender/recipient info,
 * date, subject, body, and closing.
 *
 * @module ai/templates/letter
 */

import type { DocxSchema } from '../../builder/createDocx'
import type { AiTemplate } from '../types'

/** Letter template parameters. */
export interface LetterParams {
  /** Letter body paragraphs. */
  body: string[]
  /** Recipient name. */
  recipientName: string
  /** Sender name. */
  senderName: string
  /** Closing phrase (e.g. "Sincerely", "Best regards"). */
  closing?: string
  /** Letter date. */
  date?: string
  /** Recipient address. */
  recipientAddress?: string
  /** Sender address. */
  senderAddress?: string
  /** Letter subject line. */
  subject?: string
}

const systemPrompt = `You are a professional document generator. Generate a docx-kit JSON schema for a formal letter document.

The letter should include:
1. Sender and recipient information
2. Date
3. Subject line (optional)
4. Salutation
5. Body paragraphs
6. Closing and signature

Use the following docx-kit node types:
- { type: "paragraph", text: "..." } for body text
- { type: "plugin", name: "propertyTable", options: { ... } } for address info
- { type: "plugin", name: "signatureBlock", options: { ... } } for signature

Maintain a professional tone appropriate for the letter type.
`

const schema = {
  title: 'LetterParams',
  type: 'object',
  description:
    'Formal letter with sender/recipient info, date, subject, body, and closing',
  properties: {
    closing: { description: 'Closing phrase', type: 'string' },
    date: { description: 'Letter date', type: 'string' },
    recipientAddress: { description: 'Recipient address', type: 'string' },
    senderAddress: { description: 'Sender address', type: 'string' },
    senderName: { description: 'Sender name', required: true, type: 'string' },
    subject: { description: 'Letter subject line', type: 'string' },
    body: {
      description: 'Letter body paragraphs',
      items: { type: 'string' },
      required: true,
      type: 'array',
    },
    recipientName: {
      description: 'Recipient name',
      required: true,
      type: 'string',
    },
  },
}

function generate(params: LetterParams): DocxSchema {
  const content: any[] = [{ text: params.senderName, type: 'paragraph' }]

  if (params.senderAddress) {
    content.push({ text: params.senderAddress, type: 'paragraph' })
  }
  if (params.date) {
    content.push({ text: params.date, type: 'paragraph' })
  }
  content.push({ text: params.recipientName, type: 'paragraph' })
  if (params.recipientAddress) {
    content.push({ text: params.recipientAddress, type: 'paragraph' })
  }
  if (params.subject) {
    content.push({
      level: 2,
      text: `Subject: ${params.subject}`,
      type: 'heading',
    })
  }
  content.push({ text: `Dear ${params.recipientName},`, type: 'paragraph' })
  for (const paragraph of params.body) {
    content.push({ text: paragraph, type: 'paragraph' })
  }
  content.push({
    name: 'signatureBlock',
    type: 'plugin',
    options: {
      closing: params.closing ?? 'Sincerely',
      parties: [{ name: params.senderName, role: '' }],
    },
  })

  return { content } as any as DocxSchema
}

export const letterTemplate: AiTemplate<LetterParams> = {
  generate,
  name: 'letter',
  schema: schema as any,
  systemPrompt,
  description:
    'Formal letter with sender/recipient info, date, subject, body, and closing',
}
