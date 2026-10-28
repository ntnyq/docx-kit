import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
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

  it.each([
    ['solid', 'single', 6],
    ['dashed', 'dashed', 6],
    ['dotted', 'dotted', 6],
    ['double', 'double', 18],
  ] as const)(
    'exports a visible %s divider',
    async (style, borderStyle, size) => {
      const paragraph = dividerPlugin().render(
        { color: '4472C4', style },
        createPluginTestContext(),
      )
      if (!(paragraph instanceof Paragraph)) {
        throw new TypeError('Expected a divider paragraph')
      }
      const document = new Document({ sections: [{ children: [paragraph] }] })
      const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
      const xml = await archive.file('word/document.xml')?.async('string')
      const border = xml?.match(/<w:bottom\b[^>]*\/>/)?.[0]

      expect(border).toContain(`w:val="${borderStyle}"`)
      expect(border).toContain(`w:sz="${size}"`)
      expect(border).toContain('w:color="4472C4"')
    },
  )
})
