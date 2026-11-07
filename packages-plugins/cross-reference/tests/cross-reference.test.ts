import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { crossReferencePlugin } from '../src'
import type { CrossReferenceOptions } from '../src'

async function renderXml(options: CrossReferenceOptions) {
  const paragraph = await crossReferencePlugin().render(
    options,
    createPluginTestContext(),
  )
  const archive = await JSZip.loadAsync(
    await Packer.toBuffer(
      new Document({ sections: [{ children: [paragraph] }] }),
    ),
  )
  return archive.file('word/document.xml')!.async('string')
}

describe('crossReferencePlugin', () => {
  it.each([
    { instruction: 'REF chapter \\h', kind: 'text' },
    { instruction: 'REF chapter \\w \\h', kind: 'number' },
    { instruction: 'PAGEREF chapter \\h', kind: 'page' },
  ] as const)(
    'writes the $kind field with a cached result',
    async ({ instruction, kind }) => {
      const xml = await renderXml({
        fallback: 'Chapter 1',
        kind,
        prefix: 'See ',
        suffix: '.',
        target: 'chapter',
      })
      expect(xml).toContain(`>${instruction}</w:instrText>`)
      expect(xml.match(/w:fldCharType=/g)).toHaveLength(3)
      expect(xml).toContain('w:fldCharType="begin"')
      expect(xml).toContain('w:fldCharType="separate"')
      expect(xml).toContain('w:fldCharType="end"')
      expect(xml).toContain('See ')
      expect(xml).toContain('Chapter 1')
      expect(xml).not.toContain('bookmarkStart')
      expect(xml).not.toContain('<w:hyperlink')
    },
  )

  it('defaults to text and the target as fallback, and allows disabling hyperlinks', async () => {
    const xml = await renderXml({ hyperlink: false, target: 'chapter' })
    expect(xml).toContain('>REF chapter</w:instrText>')
    expect(xml).toContain('>chapter</w:t>')
  })

  it('retains target bookmarks, field content and the part style cascade', async () => {
    const doc = createDocx({ theme: { colors: { accent: '#00AA00' } } })
      .use(
        crossReferencePlugin({
          styles: { text: { color: '#FF0000', fontWeight: 'bold' } },
        }),
      )
      .bookmark('chapter', ['Introduction'])
      .plugin(
        'crossReference',
        {
          fallback: 'Introduction',
          target: 'chapter',
          styles: {
            paragraph: { marginBottom: 0 },
            text: { color: '$colors.accent' },
          },
        },
        { color: '#0000FF', fontSize: 16, textAlign: 'right' },
      )
    const archive = await JSZip.loadAsync(await doc.toBuffer())
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml).toContain('w:name="chapter"')
    expect(xml).toContain('>REF chapter \\h</w:instrText>')
    expect(xml).toContain('<w:color w:val="00AA00"/>')
    expect(xml).toContain('<w:sz w:val="32"/>')
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('w:after="0"')
  })

  it.each([
    '',
    'bad target',
    'chapter\n',
    'chapter\r',
    'x" \\p',
    '1start',
    'a'.repeat(41),
  ])('rejects an unsafe target %s', target => {
    expect(() =>
      crossReferencePlugin().render({ target }, createPluginTestContext()),
    ).toThrow('valid Word bookmark')
  })

  it('accepts the 40-character boundary and escapes cached text', async () => {
    expect(
      await renderXml({
        fallback: '<Overview & notes>',
        target: 'a'.repeat(40),
      }),
    ).toContain('&lt;Overview &amp; notes&gt;')
  })

  it('wraps rendering errors at the compiler boundary', async () => {
    await expect(
      createDocx()
        .use(crossReferencePlugin())
        .plugin('crossReference', { target: '' })
        .toBuffer(),
    ).rejects.toMatchObject({ code: 'PLUGIN_RENDER_FAILED' })
  })
})
