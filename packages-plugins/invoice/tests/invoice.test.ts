import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { invoicePlugin } from '../src'

describe('invoicePlugin', () => {
  it.each([
    [1, 1.005, '1.01'],
    [0.7, 0.05, '0.04'],
    [1, 30_000_000_000_000, '30,000,000,000,000.00'],
    [1, 10_000_000_000_000.004, '10,000,000,000,000.00'],
    [100, 450_359_962_737.0487, '45,035,996,273,704.87'],
    [1e-7, 1e7, '1.00'],
  ])(
    'rounds quantity %s at unit price %s consistently',
    (quantity, unitPrice, expected) => {
      const result = invoicePlugin().render(
        {
          date: '2026-10-30',
          from: { name: 'Issuer' },
          invoiceNumber: 'ROUND',
          items: [{ description: 'Fractional item', quantity, unitPrice }],
          to: { name: 'Client' },
        },
        createPluginTestContext(),
      )
      expect(JSON.stringify(result)).toContain(`Subtotal: ${expected}`)
      expect(JSON.stringify(result)).toContain(`Total: ${expected}`)
    },
  )

  it('sums rounded line amounts and adds the rounded tax', () => {
    const result = invoicePlugin().render(
      {
        date: '2026-10-30',
        from: { name: 'Issuer' },
        invoiceNumber: 'ROUND',
        taxRate: 0.1,
        to: { name: 'Client' },
        items: [
          { description: 'A', quantity: 1, unitPrice: 0.335 },
          { description: 'B', quantity: 1, unitPrice: 0.335 },
        ],
      },
      createPluginTestContext(),
    )
    const text = JSON.stringify(result)
    expect(text).toContain('Subtotal: 0.68')
    expect(text).toContain('Tax: 0.07')
    expect(text).toContain('Total: 0.75')
  })

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
    if (!Array.isArray(result)) {
      throw new TypeError('Expected array result')
    }
    expect(result[0]).toBeInstanceOf(Paragraph)
    expect(result.some(item => item instanceof Table)).toBe(true)
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

  it('rejects negative line item quantities', () => {
    expect(() =>
      invoicePlugin().render(
        {
          date: '2026-06-14',
          from: { name: 'Acme Corp' },
          invoiceNumber: 'INV-003',
          items: [{ description: 'Consulting', quantity: -1, unitPrice: 5000 }],
          to: { name: 'Client Inc' },
        },
        createPluginTestContext(),
      ),
    ).toThrow(/quantity/i)
  })

  it('rejects out-of-range tax rates', () => {
    expect(() =>
      invoicePlugin().render(
        {
          date: '2026-06-14',
          from: { name: 'Acme Corp' },
          invoiceNumber: 'INV-004',
          items: [{ description: 'Consulting', quantity: 1, unitPrice: 5000 }],
          taxRate: 1.5,
          to: { name: 'Client Inc' },
        },
        createPluginTestContext(),
      ),
    ).toThrow(/tax/i)
  })
})
