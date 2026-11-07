import { createDocx, renderDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { bibliographyPlugin } from '../src'
import type { BibliographyEntry } from '../src'

const entries: BibliographyEntry[] = [
  {
    authors: ['Ada', 'Charles'],
    doi: '10.123/notes',
    id: 'ada',
    publisher: 'Press',
    title: 'Notes on computation',
    url: 'https://example.com',
    year: '1843',
  },
  { id: 'grace', title: 'Compiler design' },
]

async function xmlOf(doc: ReturnType<typeof createDocx>) {
  const archive = await JSZip.loadAsync(await doc.toBuffer())
  return archive.file('word/document.xml')!.async('string')
}

describe('bibliographyPlugin', () => {
  it.each([
    { label: 'missing ids', options: { mode: 'citation' } },
    { label: 'null ids', options: { ids: null, mode: 'citation' } },
    { label: 'string ids', options: { ids: 'ada', mode: 'citation' } },
    { label: 'number ids', options: { ids: 1, mode: 'citation' } },
    { label: 'object ids', options: { ids: {}, mode: 'citation' } },
    { label: 'non-string ID', options: { ids: ['ada', 1], mode: 'citation' } },
    { label: 'empty ID', options: { ids: [''], mode: 'citation' } },
  ])('rejects $label through the JSON compiler', async ({ options }) => {
    const document = await renderDocx({
      content: [{ name: 'bibliography', options, type: 'plugin' }],
      plugins: [{ plugin: bibliographyPlugin({ entries }), type: 'inline' }],
    })
    await expect(document.toBuffer()).rejects.toMatchObject({
      code: 'PLUGIN_RENDER_FAILED',
      cause: expect.objectContaining({
        message:
          'bibliography: citation ids must be an array of nonempty strings',
      }),
    })
  })

  it('omits only the citation when JSON explicitly supplies an empty IDs array', async () => {
    const document = await renderDocx({
      plugins: [{ plugin: bibliographyPlugin({ entries }), type: 'inline' }],
      content: [
        { text: 'Before citation', type: 'paragraph' },
        {
          name: 'bibliography',
          options: { ids: [], mode: 'citation', prefix: 'See ', suffix: '.' },
          type: 'plugin',
        },
        { text: 'After citation', type: 'paragraph' },
      ],
    })
    const xml = await xmlOf(document)
    expect(xml).toContain('Before citation')
    expect(xml).toContain('After citation')
    expect(xml).not.toContain('See ')
    expect(xml.match(/<w:p>/g)).toHaveLength(2)
  })

  it('formats numeric entries in catalog order with bookmarks and hanging indents', async () => {
    const xml = await xmlOf(
      createDocx()
        .use(bibliographyPlugin())
        .plugin('bibliography', { entries }),
    )
    expect(xml).toContain(
      '[1] Ada, Charles. Notes on computation. 1843. Press. DOI: 10.123/notes. https://example.com',
    )
    expect(xml).toContain('[2] Compiler design')
    expect(xml).toContain('w:name="bib_ada"')
    expect(xml).toContain('w:name="bib_grace"')
    expect(xml).toContain('w:hanging="420"')
    expect(xml).toContain('References')
  })

  it('reuses numbers, deduplicates IDs and permits citations before the list on repeated export', async () => {
    const doc = createDocx()
      .use(bibliographyPlugin({ entries }))
      .plugin('bibliography', {
        ids: ['grace', 'ada', 'grace'],
        mode: 'citation',
        prefix: 'See ',
        suffix: '.',
      })
      .plugin('bibliography', {
        hyperlink: false,
        ids: ['ada'],
        mode: 'citation',
      })
      .plugin('bibliography', { title: '' })
    for (let i = 0; i < 2; i++) {
      const xml = await xmlOf(doc)
      expect(xml.match(/w:anchor="bib_grace"/g)).toHaveLength(1)
      expect(xml.match(/w:anchor="bib_ada"/g)).toHaveLength(1)
      expect(xml).toMatch(/w:anchor="bib_grace"[\s\S]*?>2<\/w:t>/)
      expect(xml.indexOf('w:anchor="bib_grace"')).toBeLessThan(
        xml.indexOf('w:name="bib_grace"'),
      )
      expect(xml).not.toContain('References')
    }
  })

  it('uses invocation catalogs including empty arrays, without leaking sources between documents', async () => {
    const plugin = bibliographyPlugin({ entries })
    expect(
      await plugin.render({ entries: [] }, createPluginTestContext()),
    ).toEqual([])
    expect(
      await plugin.render(
        { ids: [], mode: 'citation' },
        createPluginTestContext(),
      ),
    ).toEqual([])
    expect(
      await bibliographyPlugin().render({}, createPluginTestContext()),
    ).toEqual([])
    const xml = await xmlOf(
      createDocx()
        .use(plugin)
        .plugin('bibliography', {
          bookmarkPrefix: 'research_',
          entries: [{ id: 'other', title: 'Different catalog' }],
        }),
    )
    expect(xml).toContain('[1] Different catalog')
    expect(xml).toContain('w:name="research_other"')
    expect(xml).not.toContain('Notes on computation')
  })

  it('applies theme and invocation styles to source and citation parts', async () => {
    const xml = await xmlOf(
      createDocx({ theme: { colors: { primary: '#123456' } } })
        .use(
          bibliographyPlugin({
            entries,
            styles: { entry: { fontWeight: 'bold' } },
          }),
        )
        .plugin('bibliography', {
          styles: {
            entry: { color: '#ABCDEF' },
            entryParagraph: { marginBottom: 0 },
          },
        })
        .plugin('bibliography', {
          ids: ['ada'],
          mode: 'citation',
          styles: {
            citation: { fontSize: 18 },
            citationParagraph: { textAlign: 'right' },
          },
        }),
    )
    expect(xml).toContain('<w:color w:val="123456"/>')
    expect(xml).toContain('<w:color w:val="ABCDEF"/>')
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:sz w:val="36"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('w:after="0"')
  })

  it('rejects unknown citations, duplicate IDs, blank titles and invalid bookmark names', () => {
    const context = createPluginTestContext()
    expect(() =>
      bibliographyPlugin({ entries }).render(
        { ids: ['missing'], mode: 'citation' },
        context,
      ),
    ).toThrow('unknown source id')
    expect(() =>
      bibliographyPlugin().render(
        { entries: [entries[0], entries[0]] },
        context,
      ),
    ).toThrow('duplicate source id')
    expect(() =>
      bibliographyPlugin().render(
        { entries: [{ id: 'x', title: '  ' }] },
        context,
      ),
    ).toThrow('title must not be empty')
    expect(() =>
      bibliographyPlugin().render(
        { entries: [{ id: 'bad id', title: 'Title' }] },
        context,
      ),
    ).toThrow('valid Word bookmark')
    expect(() =>
      bibliographyPlugin().render(
        { bookmarkPrefix: 'x'.repeat(40), entries },
        context,
      ),
    ).toThrow('valid Word bookmark')
  })
})
