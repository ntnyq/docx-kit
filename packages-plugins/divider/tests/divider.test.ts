import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { dividerPlugin } from '../src'

describe('dividerPlugin', () => {
  it('returns a plugin named "divider"', () => {
    expect(dividerPlugin().name).toBe('divider')
  })

  it('renders a solid divider', () => {
    const result = dividerPlugin().render({}, createPluginTestContext())
    expect(result).toBeInstanceOf(Paragraph)
    expect(JSON.stringify(result)).toContain('single')
  })

  it('renders a dashed divider with custom spacing and color', () => {
    const result = dividerPlugin().render(
      {
        color: '4472C4',
        spacingAfter: 240,
        spacingBefore: 120,
        style: 'dashed',
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
    const xml = JSON.stringify(result)
    expect(xml).toContain('4472C4')
    expect(xml).toContain('dashed')
  })
})
