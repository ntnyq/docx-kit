import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { pageNumberPlugin } from '../src'

describe('pageNumberPlugin', () => {
  it('returns a plugin named "pageNumber"', () => {
    expect(pageNumberPlugin().name).toBe('pageNumber')
  })

  it('renders a Paragraph', () => {
    const result = pageNumberPlugin().render({}, createPluginTestContext())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with "Page X of Y"', () => {
    const result = pageNumberPlugin().render(
      { showTotal: true },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with left alignment', () => {
    const result = pageNumberPlugin().render(
      { alignment: 'left' as any },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with custom font size', () => {
    const result = pageNumberPlugin().render(
      { fontSize: 28 },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })
})
