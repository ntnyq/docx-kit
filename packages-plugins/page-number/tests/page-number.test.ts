import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Footer, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { pageNumberPlugin } from '../src'
import type { PageNumberOptions } from '../src'

async function renderFooterXml(options: PageNumberOptions = {}) {
  const paragraph = pageNumberPlugin().render(
    options,
    createPluginTestContext(),
  )
  if (!(paragraph instanceof Paragraph)) {
    throw new TypeError('Expected a page number paragraph')
  }
  const document = new Document({
    sections: [
      {
        children: [],
        footers: { default: new Footer({ children: [paragraph] }) },
      },
    ],
  })
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
  const xml = await archive.file('word/footer1.xml')?.async('string')
  if (!xml) {
    throw new Error('Missing footer XML')
  }
  return xml
}

describe('pageNumberPlugin', () => {
  it('returns a plugin named "pageNumber"', () => {
    expect(pageNumberPlugin().name).toBe('pageNumber')
  })

  it('renders a Paragraph with no options', () => {
    const result = pageNumberPlugin().render({}, createPluginTestContext())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders a Paragraph with showTotal: true', () => {
    const result = pageNumberPlugin().render(
      { showTotal: true },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders a Paragraph with custom alignment and fontSize', () => {
    const result = pageNumberPlugin().render(
      { alignment: 'left', fontSize: 12 },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('emits current and total page fields in the footer', async () => {
    const xml = await renderFooterXml({ showTotal: true })
    expect(xml).toMatch(/<w:instrText[^>]*>\s*PAGE\s*<\/w:instrText>/)
    expect(xml).toMatch(/<w:instrText[^>]*>\s*NUMPAGES\s*<\/w:instrText>/)
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2)
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(2)
    expect(xml).toContain('>Page </w:t>')
    expect(xml).toContain('> of </w:t>')
  })

  it('emits only the current page field by default', async () => {
    const xml = await renderFooterXml()
    expect(xml).toMatch(/<w:instrText[^>]*>\s*PAGE\s*<\/w:instrText>/)
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1)
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1)
    expect(xml).not.toContain('NUMPAGES')
  })

  it('converts user fontSize (points) to docx size (half-points)', async () => {
    const xml = await renderFooterXml({ fontSize: 14 })
    expect(xml).toContain('<w:sz w:val="28"/>')
  })
})
