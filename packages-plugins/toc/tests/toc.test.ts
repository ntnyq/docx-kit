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
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
    expect((result as any[])[1]).toBeInstanceOf(TableOfContents)
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
})
