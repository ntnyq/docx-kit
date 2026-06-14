import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, TableOfContents } from 'docx'
import { describe, expect, it } from 'vitest'
import { tocPlugin } from '../src'

describe('tocPlugin', () => {
  it('returns a plugin named "toc"', () => {
    expect(tocPlugin().name).toBe('toc')
  })

  it('renders a title and table of contents field', () => {
    const result = tocPlugin().render(
      { maxLevel: 3, title: 'Contents' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    if (!Array.isArray(result)) {
      throw new TypeError('Expected array result')
    }
    expect(result[0]).toBeInstanceOf(Paragraph)
    expect(result[1]).toBeInstanceOf(TableOfContents)
  })

  it('uses a custom title and heading range', () => {
    const result = tocPlugin().render(
      { maxLevel: 2, title: '目录' },
      createPluginTestContext(),
    )
    const xml = JSON.stringify(result)
    expect(xml).toContain('目录')
    expect(xml).toContain('1-2')
  })

  it('does not render the title as an indexed heading', () => {
    const result = tocPlugin().render(
      { maxLevel: 3, title: 'Contents' },
      createPluginTestContext(),
    )
    if (!Array.isArray(result)) {
      throw new TypeError('Expected array result')
    }
    const xml = JSON.stringify(result[0])
    expect(xml).not.toContain('Heading1')
  })

  it('clamps maxLevel to a valid heading range', () => {
    const result = tocPlugin().render(
      { maxLevel: 0, title: 'Contents' },
      createPluginTestContext(),
    )
    const xml = JSON.stringify(result)
    expect(xml).toContain('1-1')
    expect(xml).not.toContain('1-0')
  })
})
