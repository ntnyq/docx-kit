import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { codeBlockPlugin } from '../../../packages-plugins/code-block/src'
import { createDocx, defineStyles } from '../../core/src'
import { markdownToNodes as umbrellaMarkdownToNodes } from '../../docx-kit/src/markdown'
import { markdownToNodes } from '../src'

describe('Markdown document integration', () => {
  it('exposes the converter through the umbrella subpath', () => {
    expect(umbrellaMarkdownToNodes('# Title')).toEqual(
      markdownToNodes('# Title'),
    )
  })

  it('compiles Markdown using existing styles, theme tokens and plugins', async () => {
    const doc = createDocx({
      theme: { colors: { primary: '#123456' } },
      styles: defineStyles({
        body: { color: '$colors.primary', fontSize: 13 },
        h1: { color: '#654321', fontSize: 24 },
        source: { color: '#112233' },
      }),
    }).use(codeBlockPlugin())
    const source = [
      '# Markdown document',
      '',
      'Hello **bold**  ',
      '[Docs](https://example.com)',
      '',
      '- [x] Complete',
      '',
      '> Quoted body',
      '',
      '| Label | Value |',
      '| --- | ---: |',
      '| **Revenue** | [Details](https://example.com/table) |',
      '',
      '```',
      'const answer = 42',
      '```',
      '',
      '- parent',
      '  3. nested start',
      '',
      '---',
    ].join('\n')
    for (const node of markdownToNodes(source, {
      classNames: { code: 'source', paragraph: 'body' },
    })) {
      doc.add(node)
    }
    const archive = await JSZip.loadAsync(await doc.toUint8Array())
    const xml = await archive.file('word/document.xml')?.async('string')
    const relations = await archive
      .file('word/_rels/document.xml.rels')
      ?.async('string')
    const numbering = await archive.file('word/numbering.xml')?.async('string')

    expect(xml).toContain('Markdown document')
    expect(xml).toContain('w:val="654321"')
    expect(xml).toContain('w:val="123456"')
    expect(xml).toContain('w:val="112233"')
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:br/>')
    expect(xml).toContain('<w14:checkbox>')
    expect(xml).toContain('<w:tbl>')
    expect(xml).toContain('Revenue')
    expect(xml).toContain('Details')
    expect(xml).toContain('const answer = 42')
    expect(xml).toContain('<w:ilvl w:val="1"/>')
    expect(xml).not.toContain('[object Object]')
    expect(relations).toContain('https://example.com/table')
    expect(numbering).toMatch(/<w:lvl w:ilvl="1"[^>]*><w:start w:val="3"\/>/)
  })

  it('reports missing plugins through the existing compiler', async () => {
    const doc = createDocx()
    for (const node of markdownToNodes('```\ncode\n```')) {
      doc.add(node)
    }
    await expect(doc.toUint8Array()).rejects.toThrow(/codeBlock/)
  })

  it('renders filtered and sorted table rows from a transform with their rich content', async () => {
    const nodes = markdownToNodes(
      [
        '| Name |',
        '| --- |',
        '| **Alpha** |',
        '| *Beta* |',
        '| [Gamma](https://example.com/gamma) |',
      ].join('\n'),
      {
        transform(nodes, context) {
          if (context.type === 'table') {
            for (const node of nodes) {
              if (node.type === 'table') {
                node.data = node.data
                  .filter(row => row.column0 !== 'Alpha')
                  .sort((left, right) =>
                    String(right.column0).localeCompare(String(left.column0)),
                  )
              }
            }
          }
          return nodes
        },
      },
    )
    const doc = createDocx()
    for (const node of nodes) {
      doc.add(node)
    }
    const archive = await JSZip.loadAsync(await doc.toUint8Array())
    const xml = await archive.file('word/document.xml')?.async('string')
    const relations = await archive
      .file('word/_rels/document.xml.rels')
      ?.async('string')

    expect(xml).toBeDefined()
    expect(
      [...(xml?.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g) ?? [])].map(
        match => match[1],
      ),
    ).toEqual(['Name', 'Gamma', 'Beta'])
    expect(xml).toContain('<w:i/>')
    expect(xml).toContain('<w:hyperlink')
    expect(relations).toContain('https://example.com/gamma')
  })

  it('renders edited and appended table values without losing unchanged formatting', async () => {
    const [node] = markdownToNodes(
      '| Name | Value |\n| --- | --- |\n| **Alpha** | 10 |',
    )
    if (node.type !== 'table') {
      throw new Error('Expected a table')
    }
    node.data[0].column1 = '42'
    node.data.push({ column0: 'Delta', column1: 0 })
    const archive = await JSZip.loadAsync(
      await createDocx().add(node).toUint8Array(),
    )
    const xml = await archive.file('word/document.xml')?.async('string')

    expect(xml).toBeDefined()
    expect(
      [...(xml?.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g) ?? [])].map(
        match => match[1],
      ),
    ).toEqual(['Name', 'Value', 'Alpha', '42', 'Delta', '0'])
    expect(xml).toMatch(/<w:b\/><w:bCs\/><\/w:rPr><w:t[^>]*>Alpha<\/w:t>/)
  })

  it.each([
    { blockquote: 'quote muted', paragraph: 'body lead' },
    { blockquote: ['quote', 'muted'], paragraph: ['body', 'lead'] },
    { blockquote: ['quote muted'], paragraph: ['body lead'] },
  ])(
    'compiles quote class combinations with the existing stylesheet: %j',
    async classNames => {
      const doc = createDocx({
        styles: defineStyles({
          body: { color: '#112233' },
          lead: { fontSize: 14 },
          muted: { color: '#445566' },
          quote: { italic: true },
        }),
      })
      for (const node of markdownToNodes('> Quote', { classNames })) {
        doc.add(node)
      }
      const archive = await JSZip.loadAsync(await doc.toUint8Array())
      const xml = await archive.file('word/document.xml')?.async('string')

      expect(xml).toContain('Quote')
      expect(xml).toContain('w:val="445566"')
      expect(xml).toContain('<w:sz w:val="28"/>')
      expect(xml).toContain('<w:i/>')
    },
  )
})
