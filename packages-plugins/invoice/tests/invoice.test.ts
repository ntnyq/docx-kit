import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { invoicePlugin } from '../src'

describe('invoicePlugin', () => {
  it('returns a plugin named "invoice"', () => {
    expect(invoicePlugin().name).toBe('invoice')
  })

  it('renders invoice header, parties, item table, and totals', () => {
    const result = invoicePlugin().render(
      {
        date: '2026-06-14',
        from: { name: 'Acme Corp' },
        invoiceNumber: 'INV-001',
        items: [{ description: 'Consulting', quantity: 2, unitPrice: 5000 }],
        taxRate: 0.1,
        to: { name: 'Client Inc' },
      },
      createPluginTestContext(),
    )

    expect(Array.isArray(result)).toBe(true)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
    expect((result as any[]).some(item => item instanceof Table)).toBe(true)
    expect(JSON.stringify(result)).toContain('INV-001')
    expect(JSON.stringify(result)).toContain('11,000.00')
  })

  it('renders notes and due date when provided', () => {
    const result = invoicePlugin().render(
      {
        currency: 'USD',
        date: '2026-06-14',
        dueDate: '2026-06-30',
        from: { name: 'Acme Corp' },
        invoiceNumber: 'INV-002',
        items: [{ description: 'Design', quantity: 1, unitPrice: 1200 }],
        notes: 'Thank you for your business.',
        to: { name: 'Client Inc' },
      },
      createPluginTestContext(),
    )

    expect(Array.isArray(result)).toBe(true)
    const xml = JSON.stringify(result)
    expect(xml).toContain('2026-06-30')
    expect(xml).toContain('Thank you for your business.')
    expect(xml).toContain('1,200.00')
  })
})
