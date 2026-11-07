import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { glossaryPlugin } from '../src'
import type { GlossaryOptions } from '../src'

async function renderXml(options: GlossaryOptions) {
  const archive = await JSZip.loadAsync(
    await createDocx()
      .use(glossaryPlugin())
      .plugin('glossary', options)
      .toBuffer(),
  )
  return archive.file('word/document.xml')!.async('string')
}

describe('glossaryPlugin', () => {
  it('sorts without mutating input and includes full forms and repeated table headers', async () => {
    const entries = [
      { definition: 'Archive format', term: 'ZIP' },
      {
        definition: 'Software interface',
        fullForm: 'Application Programming Interface',
        term: 'API',
      },
    ]
    const xml = await renderXml({ entries })
    expect(xml.indexOf('>API</w:t>')).toBeLessThan(xml.indexOf('>ZIP</w:t>'))
    expect(entries[0].term).toBe('ZIP')
    expect(xml).toContain('Application Programming Interface')
    expect(xml).toContain('<w:tblHeader/>')
    expect(xml).toContain('Glossary')
    // Long definitions can flow across pages rather than forcing oversized rows.
    expect(xml).not.toContain('cantSplit')
  })

  it('groups nonadjacent entries with unnamed entries first and customizable labels', async () => {
    const xml = await renderXml({
      groupBy: 'group',
      labels: { definition: '解释', fullForm: '全称', term: '术语' },
      title: '',
      entries: [
        { definition: 'Z', group: 'Animals', term: 'Zebra' },
        { definition: 'No group', term: 'Other' },
        { definition: 'Interface', group: 'Software', term: 'API' },
        { definition: 'A', group: 'Animals', term: 'Ant' },
      ],
    })
    expect(xml.match(/<w:tbl>/g)).toHaveLength(3)
    expect(xml.match(/>Animals<\/w:t>/g)).toHaveLength(1)
    expect(xml.indexOf('No group')).toBeLessThan(xml.indexOf('Animals'))
    expect(xml.indexOf('Animals')).toBeLessThan(xml.indexOf('Software'))
    expect(xml.indexOf('>Ant</w:t>')).toBeLessThan(xml.indexOf('>Zebra</w:t>'))
    expect(xml).toContain('术语')
    expect(xml).not.toContain('Glossary')
  })

  it('preserves term and group first-appearance order with sort none', async () => {
    const xml = await renderXml({
      groupBy: 'group',
      sort: 'none',
      entries: [
        { definition: 'last letter', group: 'Z group', term: 'Z' },
        { definition: 'second', group: 'A group', term: 'B' },
        { definition: 'first', group: 'Z group', term: 'A' },
      ],
    })
    expect(xml.indexOf('Z group')).toBeLessThan(xml.indexOf('A group'))
    expect(xml.indexOf('>Z</w:t>')).toBeLessThan(xml.indexOf('>A</w:t>'))
  })

  it('supports locale collation, long text and duplicate terms', async () => {
    const xml = await renderXml({
      locale: 'sv',
      entries: [
        { definition: 'Swedish last letters', term: 'Ä' },
        { definition: 'Long '.repeat(500), term: 'Z' },
        { definition: 'Second meaning', term: 'Z' },
      ],
    })
    expect(xml.indexOf('>Z</w:t>')).toBeLessThan(xml.indexOf('>Ä</w:t>'))
    expect(xml.match(/>Z<\/w:t>/g)).toHaveLength(2)
    expect(xml).toContain('Long '.repeat(500))
  })

  it('applies default, theme and invocation part styles', async () => {
    const doc = createDocx({
      theme: { colors: { primary: '#123456', surface: '#EEEEEE' } },
    })
      .use(glossaryPlugin({ styles: { term: { color: '#FF0000' } } }))
      .plugin('glossary', {
        entries: [{ definition: 'Interface', term: 'API' }],
        styles: {
          bodyCell: { marginLeft: 0 },
          bodyParagraph: { textAlign: 'right' },
          term: { color: '#00AA00' },
        },
      })
    const archive = await JSZip.loadAsync(await doc.toBuffer())
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:color w:val="123456"/>')
    expect(xml).toContain('w:fill="EEEEEE"')
    expect(xml).toContain('<w:color w:val="00AA00"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('w:w="0"')
  })

  it('omits empty tables and rejects blank terms or definitions', async () => {
    const context = createPluginTestContext()
    expect(await glossaryPlugin().render({ entries: [] }, context)).toEqual([])
    for (const entry of [
      { definition: 'Meaning', term: '' },
      { definition: '  ', term: 'API' },
    ]) {
      expect(() =>
        glossaryPlugin().render({ entries: [entry] }, context),
      ).toThrow('must not be empty')
    }
  })
})
