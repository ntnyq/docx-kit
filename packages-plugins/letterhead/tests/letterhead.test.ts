import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { letterheadPlugin } from '../src'

describe('letterheadPlugin', () => {
  it('returns a plugin named "letterhead"', () => {
    expect(letterheadPlugin().name).toBe('letterhead')
  })

  it('renders company identity and contact info', () => {
    const result = letterheadPlugin().render(
      {
        companyName: 'Acme Corp',
        email: 'info@acme.com',
        phone: '+1 555 0100',
        tagline: 'Innovation First',
      },
      createPluginTestContext(),
    )

    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).every(item => item instanceof Paragraph)).toBe(
      true,
    )
    const xml = JSON.stringify(result)
    expect(xml).toContain('Acme Corp')
    expect(xml).toContain('info@acme.com')
  })

  it('renders address and website when provided', () => {
    const result = letterheadPlugin().render(
      {
        address: '1 Infinite Loop',
        companyName: 'Acme Corp',
        website: 'https://acme.test',
      },
      createPluginTestContext(),
    )
    const xml = JSON.stringify(result)
    expect(xml).toContain('1 Infinite Loop')
    expect(xml).toContain('https://acme.test')
  })
})
