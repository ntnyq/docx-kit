import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { badgePlugin } from '../src'

describe('badgePlugin', () => {
  it('returns a plugin named "badge"', () => {
    expect(badgePlugin().name).toBe('badge')
  })

  it('renders a badge paragraph with preset color', () => {
    const result = badgePlugin().render(
      { color: 'warning', text: 'DRAFT' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
    expect(JSON.stringify(result)).toContain('DRAFT')
    expect(JSON.stringify(result)).toContain('FFF4D6')
  })

  it('renders a badge paragraph with custom colors', () => {
    const result = badgePlugin().render(
      {
        backgroundColor: '111111',
        color: 'FFFFFF',
        text: 'APPROVED',
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
    const xml = JSON.stringify(result)
    expect(xml).toContain('APPROVED')
    expect(xml).toContain('111111')
    expect(xml).toContain('FFFFFF')
  })
})
