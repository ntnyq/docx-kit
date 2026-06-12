import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { coverPagePlugin } from '../src'

describe('coverPagePlugin', () => {
  it('returns a plugin named "coverPage"', () => {
    expect(coverPagePlugin().name).toBe('coverPage')
  })

  it('renders basic cover page with title only', () => {
    const result = coverPagePlugin().render(
      { title: 'Annual Report' },
      createPluginTestContext(),
    ) as Paragraph[]
    expect(Array.isArray(result)).toBe(true)
    expect(result.length).toBeGreaterThan(0)
    expect(result[0]).toBeInstanceOf(Paragraph)
  })

  it('renders cover page with all fields', () => {
    const result = coverPagePlugin().render(
      {
        author: 'John Doe',
        date: '2026-06-11',
        organization: 'Acme Corp',
        subtitle: '2026 Edition',
        title: 'Full Report',
      },
      createPluginTestContext(),
    ) as Paragraph[]
    expect(Array.isArray(result)).toBe(true)
    for (const p of result) {
      expect(p).toBeInstanceOf(Paragraph)
    }
  })

  it('hides rule when showRule is false', () => {
    const resultWithRule = coverPagePlugin().render(
      { showRule: true, title: 'Test' },
      createPluginTestContext(),
    ) as Paragraph[]
    const resultWithoutRule = coverPagePlugin().render(
      { showRule: false, title: 'Test' },
      createPluginTestContext(),
    ) as Paragraph[]
    expect(resultWithRule.length).toBeGreaterThan(resultWithoutRule.length)
  })

  it('renders with background color', () => {
    const result = coverPagePlugin().render(
      { backgroundColor: 'F0F0F0', title: 'Styled' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
  })

  it('renders with right alignment', () => {
    const result = coverPagePlugin().render(
      { alignment: 'right' as any, title: 'Aligned' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
  })

  it('renders with subtitle and date only', () => {
    const result = coverPagePlugin().render(
      { date: '2026-06-11', subtitle: 'Summary', title: 'Simple' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
  })
})
