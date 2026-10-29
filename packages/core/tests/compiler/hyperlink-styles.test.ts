import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { createDocx } from '../../src'

describe('block hyperlink paragraph styles', () => {
  it.each(['external', 'internal'] as const)(
    'applies defaults, classes and overrides to %s links',
    async kind => {
      const doc = createDocx({
        theme: { colors: { primary: '#123456' } },
        defaults: {
          paragraph: { keepNext: true, marginTop: 12, textAlign: 'right' },
        },
        styles: {
          link: {
            color: '$colors.primary',
            marginTop: 18,
            textAlign: 'center',
          },
        },
      })
      const options = { className: 'link' as const, style: { marginTop: 24 } }
      if (kind === 'external') {
        doc.hyperlink('https://example.com', 'Link', options)
      } else {
        doc.internalLink('target', 'Link', options)
      }
      const zip = await JSZip.loadAsync(await doc.toBuffer())
      const xml = await zip.file('word/document.xml')!.async('string')
      expect(xml).toContain('<w:jc w:val="center"/>')
      expect(xml).toContain('<w:spacing w:before="480"/>')
      expect(xml).toContain('<w:keepNext/>')
      expect(xml).toContain('<w:color w:val="123456"/>')
    },
  )

  it('keeps inline link layout on its containing paragraph', async () => {
    const doc = createDocx().p(
      [
        {
          children: ['Inline'],
          style: { marginTop: 24, textAlign: 'center' },
          type: 'hyperlink',
          url: 'https://example.com',
        },
      ],
      { style: { marginTop: 12, textAlign: 'right' } },
    )
    const zip = await JSZip.loadAsync(await doc.toBuffer())
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('<w:spacing w:before="240"/>')
    expect(xml).not.toContain('w:val="center"')
  })
})
