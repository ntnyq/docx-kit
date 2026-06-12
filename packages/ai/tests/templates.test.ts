/**
 * Tests for AI template definitions.
 *
 * @module tests/ai/templates
 */

import { describe, expect, it } from 'vitest'
import {
  invoiceTemplate,
  letterTemplate,
  reportTemplate,
  resumeTemplate,
} from '../src/templates'
import type { AiTemplate } from '../src/types'

function validateTemplate(template: AiTemplate<any>): void {
  expect(template.name).toBeTypeOf('string')
  expect(template.name.length).toBeGreaterThan(0)
  expect(template.description).toBeTypeOf('string')
  expect(template.systemPrompt).toBeTypeOf('string')
  expect(template.systemPrompt.length).toBeGreaterThan(0)
  expect(template.schema).toBeTypeOf('object')
  expect(template.schema.type).toBe('object')
  expect(template.schema.properties).toBeTypeOf('object')
  expect(template.generate).toBeTypeOf('function')
}

describe('report template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(reportTemplate as any)
  })

  it('generates schema with title', () => {
    const result = reportTemplate.generate({ title: 'Annual Report' })
    expect(result.content).toBeInstanceOf(Array)
    expect(result.content.length).toBeGreaterThan(0)
  })

  it('generates cover page plugin node', () => {
    const result = reportTemplate.generate({ author: 'Alice', title: 'Report' })
    const coverNode = result.content.find(
      (n: any) => n.type === 'plugin' && n.name === 'coverPage',
    )
    expect(coverNode).toBeDefined()
  })

  it('generates executive summary section', () => {
    const result = reportTemplate.generate({
      executiveSummary: ['Key finding 1', 'Key finding 2'],
      title: 'Report',
    })
    const headingNode = result.content.find(
      (n: any) => n.type === 'heading' && n.text === 'Executive Summary',
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
      (n: any) => n.type === 'heading' && n.level === 2,
    )
    expect(sectionHeadings.length).toBeGreaterThanOrEqual(2)
  })

  it('generates conclusion', () => {
    const result = reportTemplate.generate({
      conclusion: 'In summary, the project was successful.',
      title: 'Report',
    })
    const conclusionNode = result.content.find(
      (n: any) => n.type === 'heading' && n.text === 'Conclusion',
    )
    expect(conclusionNode).toBeDefined()
  })

  it('schema has title as required property', () => {
    expect(reportTemplate.schema.properties.title.required).toBe(true)
  })
})

describe('invoice template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(invoiceTemplate as any)
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
    const tableNode = result.content.find((n: any) => n.type === 'table')
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
    const totalNode = result.content.find(
      (n: any) =>
        n.type === 'paragraph'
        && typeof n.text === 'string'
        && n.text.startsWith('Total:'),
    )
    expect(totalNode).toBeDefined()
    expect((totalNode as any)!.text).toContain('385.00')
  })
})

describe('resume template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(resumeTemplate as any)
  })

  it('generates schema with name', () => {
    const result = resumeTemplate.generate({ name: 'John Doe' })
    expect(result.content).toBeInstanceOf(Array)
    const nameHeading = result.content.find(
      (n: any) => n.type === 'heading' && n.text === 'John Doe',
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
      (n: any) => n.type === 'heading' && n.text === 'Experience',
    )
    expect(expHeading).toBeDefined()
    const bulletList = result.content.find((n: any) => n.type === 'bulletList')
    expect(bulletList).toBeDefined()
  })

  it('generates education section', () => {
    const result = resumeTemplate.generate({
      education: [{ degree: 'BSc CS', institution: 'MIT', year: '2018' }],
      name: 'John Doe',
    })
    const eduHeading = result.content.find(
      (n: any) => n.type === 'heading' && n.text === 'Education',
    )
    expect(eduHeading).toBeDefined()
  })

  it('generates skills section', () => {
    const result = resumeTemplate.generate({
      name: 'John Doe',
      skills: ['TypeScript', 'React', 'Node.js'],
    })
    const skillsHeading = result.content.find(
      (n: any) => n.type === 'heading' && n.text === 'Skills',
    )
    expect(skillsHeading).toBeDefined()
    const skillsParagraph = result.content.find(
      (n: any) =>
        n.type === 'paragraph'
        && typeof n.text === 'string'
        && n.text.includes('TypeScript'),
    )
    expect(skillsParagraph).toBeDefined()
  })
})

describe('letter template', () => {
  // eslint-disable-next-line vitest/expect-expect
  it('has valid structure', () => {
    validateTemplate(letterTemplate as any)
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
      (n: any) =>
        n.type === 'paragraph'
        && typeof n.text === 'string'
        && n.text.startsWith('Dear'),
    )
    expect(salutation).toBeDefined()
  })

  it('generates signature block', () => {
    const result = letterTemplate.generate({
      body: ['Body text.'],
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    const signature = result.content.find(
      (n: any) => n.type === 'plugin' && n.name === 'signatureBlock',
    )
    expect(signature).toBeDefined()
  })

  it('uses custom closing phrase', () => {
    const result = letterTemplate.generate({
      body: ['Body text.'],
      closing: 'Best regards',
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    const signature = result.content.find(
      (n: any) => n.type === 'plugin' && n.name === 'signatureBlock',
    )
    expect(signature).toBeDefined()
    expect((signature as any).options.closing).toBe('Best regards')
  })
})
