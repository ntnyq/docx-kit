/**
 * Tests for AI template definitions.
 *
 * @module tests/ai/templates
 */

import { renderDocx } from '@docxkit/core'
import { describe, expect, it } from 'vitest'
import {
  invoiceTemplate,
  letterTemplate,
  reportTemplate,
  resumeTemplate,
} from '../src/templates'
import type {
  BlockNode,
  HeadingNode,
  ParagraphNode,
  TableNode,
} from '@docxkit/core'
import type {
  AiTemplate,
  AiTemplateObjectSchema,
  AiTemplateSchema,
} from '../src/types'

function isHeadingNode(node: BlockNode): node is HeadingNode {
  return node.type === 'heading'
}

function isObjectSchema(
  schema: AiTemplateSchema,
): schema is AiTemplateObjectSchema {
  return schema.type === 'object'
}

function isParagraphNode(node: BlockNode): node is ParagraphNode {
  return node.type === 'paragraph'
}

function isTableNode(
  node: BlockNode,
): node is TableNode<Record<string, unknown>> {
  return node.type === 'table'
}

function isTextParagraphNode(
  node: BlockNode,
): node is ParagraphNode & { text: string } {
  return node.type === 'paragraph' && typeof node.text === 'string'
}

function validateTemplate<TParams extends object>(
  template: AiTemplate<TParams>,
): void {
  expect(template.name).toBeTypeOf('string')
  expect(template.name.length).toBeGreaterThan(0)
  expect(template.description).toBeTypeOf('string')
  expect(template.systemPrompt).toBeTypeOf('string')
  expect(template.systemPrompt.length).toBeGreaterThan(0)
  expect(template.schema).toBeTypeOf('object')
  expect(template.schema.type).toBe('object')
  expect(isObjectSchema(template.schema)).toBe(true)
  expect(
    isObjectSchema(template.schema) && template.schema.properties,
  ).toBeTypeOf('object')
  expect(template.generate).toBeTypeOf('function')
}

describe('report template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(reportTemplate)
  })

  it('generates schema with title', () => {
    const result = reportTemplate.generate({ title: 'Annual Report' })
    expect(result.content).toBeInstanceOf(Array)
    expect(result.content.length).toBeGreaterThan(0)
  })

  it('generates a core-only cover page', () => {
    const result = reportTemplate.generate({ author: 'Alice', title: 'Report' })
    const title = result.content.find(
      node => isHeadingNode(node) && node.text === 'Report',
    )
    expect(title).toBeDefined()
    expect(result.content.some(node => node.type === 'pageBreak')).toBe(true)
    expect(result.content.some(node => node.type === 'plugin')).toBe(false)
  })

  it('generates executive summary section', () => {
    const result = reportTemplate.generate({
      executiveSummary: ['Key finding 1', 'Key finding 2'],
      title: 'Report',
    })
    const headingNode = result.content.find(
      node => isHeadingNode(node) && node.text === 'Executive Summary',
    )
    expect(headingNode).toBeDefined()
  })

  it('generates body sections', () => {
    const result = reportTemplate.generate({
      title: 'Report',
      sections: [
        { content: ['Paragraph A1'], heading: 'Section A' },
        { content: ['Paragraph B1'], heading: 'Section B' },
      ],
    })
    const sectionHeadings = result.content.filter(
      node => isHeadingNode(node) && node.level === 2,
    )
    expect(sectionHeadings.length).toBeGreaterThanOrEqual(2)
  })

  it('generates conclusion', () => {
    const result = reportTemplate.generate({
      conclusion: 'In summary, the project was successful.',
      title: 'Report',
    })
    const conclusionNode = result.content.find(
      node => isHeadingNode(node) && node.text === 'Conclusion',
    )
    expect(conclusionNode).toBeDefined()
  })

  it('schema has title as required property', () => {
    expect(isObjectSchema(reportTemplate.schema)).toBe(true)
    if (!isObjectSchema(reportTemplate.schema)) {
      throw new Error('report template schema must be an object schema')
    }
    expect(reportTemplate.schema.properties.title?.required).toBe(true)
  })
})

describe('invoice template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(invoiceTemplate)
  })

  it('generates schema with items', () => {
    const result = invoiceTemplate.generate({
      clientName: 'Client Inc',
      invoiceNumber: 'INV-001',
      issuerName: 'My Company',
      items: [{ description: 'Service A', quantity: 2, unitPrice: 100 }],
    })
    expect(result.content).toBeInstanceOf(Array)
    expect(result.content.length).toBeGreaterThan(0)
  })

  it('generates invoice table', () => {
    const result = invoiceTemplate.generate({
      clientName: 'Client Inc',
      invoiceNumber: 'INV-001',
      issuerName: 'My Company',
      items: [{ description: 'Service A', quantity: 2, unitPrice: 100 }],
    })
    const tableNode = result.content.find(isTableNode)
    expect(tableNode).toBeDefined()
  })

  it('calculates totals correctly', () => {
    const result = invoiceTemplate.generate({
      clientName: 'Client Inc',
      invoiceNumber: 'INV-001',
      issuerName: 'My Company',
      taxRate: 0.1,
      items: [
        { description: 'Item 1', quantity: 3, unitPrice: 50 },
        { description: 'Item 2', quantity: 1, unitPrice: 200 },
      ],
    })
    // Subtotal: 3*50 + 1*200 = 350, Tax: 35, Total: 385
    const totalNode = result.content
      .filter(isTextParagraphNode)
      .find(node => node.text.startsWith('Total:'))
    expect(totalNode).toBeDefined()
    expect(totalNode?.text ?? '').toContain('385.00')
  })
})

describe('resume template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(resumeTemplate)
  })

  it('generates schema with name', () => {
    const result = resumeTemplate.generate({ name: 'John Doe' })
    expect(result.content).toBeInstanceOf(Array)
    const nameHeading = result.content.find(
      node => isHeadingNode(node) && node.text === 'John Doe',
    )
    expect(nameHeading).toBeDefined()
  })

  it('generates experience section', () => {
    const result = resumeTemplate.generate({
      name: 'John Doe',
      experience: [
        {
          company: 'Tech Corp',
          endDate: 'Present',
          highlights: ['Built system A', 'Led team B'],
          role: 'Developer',
          startDate: '2020',
        },
      ],
    })
    const expHeading = result.content.find(
      node => isHeadingNode(node) && node.text === 'Experience',
    )
    expect(expHeading).toBeDefined()
    const bulletList = result.content.find(node => node.type === 'bulletList')
    expect(bulletList).toBeDefined()
  })

  it('generates education section', () => {
    const result = resumeTemplate.generate({
      education: [{ degree: 'BSc CS', institution: 'MIT', year: '2018' }],
      name: 'John Doe',
    })
    const eduHeading = result.content.find(
      node => isHeadingNode(node) && node.text === 'Education',
    )
    expect(eduHeading).toBeDefined()
  })

  it('generates skills section', () => {
    const result = resumeTemplate.generate({
      name: 'John Doe',
      skills: ['TypeScript', 'React', 'Node.js'],
    })
    const skillsHeading = result.content.find(
      node => isHeadingNode(node) && node.text === 'Skills',
    )
    expect(skillsHeading).toBeDefined()
    const skillsParagraph = result.content.find(
      node =>
        isParagraphNode(node)
        && typeof node.text === 'string'
        && node.text.includes('TypeScript'),
    )
    expect(skillsParagraph).toBeDefined()
  })
})

describe('letter template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(letterTemplate)
  })

  it('generates schema with sender and recipient', () => {
    const result = letterTemplate.generate({
      body: ['This is the letter body.'],
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    expect(result.content).toBeInstanceOf(Array)
    expect(result.content.length).toBeGreaterThan(0)
  })

  it('generates salutation', () => {
    const result = letterTemplate.generate({
      body: ['Body text.'],
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    const salutation = result.content.find(
      node =>
        isParagraphNode(node)
        && typeof node.text === 'string'
        && node.text.startsWith('Dear'),
    )
    expect(salutation).toBeDefined()
  })

  it('generates closing and sender paragraphs', () => {
    const result = letterTemplate.generate({
      body: ['Body text.'],
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    const closing = result.content.find(
      node => isTextParagraphNode(node) && node.text === 'Sincerely',
    )
    const sender = result.content.find(
      node => isTextParagraphNode(node) && node.text === 'Alice',
    )
    expect(closing).toBeDefined()
    expect(sender).toBeDefined()
    expect(result.content.some(node => node.type === 'plugin')).toBe(false)
  })

  it('uses custom closing phrase', () => {
    const result = letterTemplate.generate({
      body: ['Body text.'],
      closing: 'Best regards',
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    const closing = result.content.find(
      node => isTextParagraphNode(node) && node.text === 'Best regards',
    )
    expect(closing).toBeDefined()
  })
})

describe('template rendering', () => {
  it.each([
    [
      'report',
      reportTemplate.generate({
        author: 'Alice',
        executiveSummary: ['Summary'],
        title: 'Report',
      }),
    ],
    [
      'invoice',
      invoiceTemplate.generate({
        clientName: 'Client',
        invoiceNumber: 'INV-1',
        issuerName: 'Issuer',
        items: [{ description: 'Service', quantity: 1, unitPrice: 100 }],
      }),
    ],
    [
      'resume',
      resumeTemplate.generate({
        email: 'alice@example.com',
        name: 'Alice',
        skills: ['TypeScript'],
      }),
    ],
    [
      'letter',
      letterTemplate.generate({
        body: ['Hello'],
        recipientName: 'Bob',
        senderName: 'Alice',
      }),
    ],
  ])(
    'renders the %s template without plugin registration',
    async (_, schema) => {
      const document = await renderDocx(schema)
      const buffer = await document.toBuffer()

      expect(buffer.byteLength).toBeGreaterThan(0)
    },
  )
})
