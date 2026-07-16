import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import type { BlockNode, TableNode } from '@docxkit/types'

async function renderPackage(nodes: BlockNode[]) {
  const document = await compileDocument({
    config: {},
    nodes,
    plugins: new Map(),
  })
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document))

  async function read(path: string) {
    const file = archive.file(path)
    if (!file) {
      throw new Error(`Missing OOXML part: ${path}`)
    }
    return file.async('string')
  }

  return { read }
}

describe('compiler OOXML contracts', () => {
  it('preserves an inline hyperlink inside a paragraph', async () => {
    const pkg = await renderPackage([
      {
        type: 'paragraph',
        children: [
          { text: 'Visit ', type: 'text' },
          {
            children: ['docx-kit'],
            type: 'hyperlink',
            url: 'https://example.com/docx-kit',
          },
        ],
      },
    ])

    const documentXml = await pkg.read('word/document.xml')
    const relationshipsXml = await pkg.read('word/_rels/document.xml.rels')

    expect(documentXml).toContain('<w:hyperlink')
    expect(documentXml).toContain('Visit ')
    expect(documentXml).toContain('docx-kit')
    expect(relationshipsXml).toContain('https://example.com/docx-kit')
    expect(relationshipsXml).toContain('relationships/hyperlink')
  })

  it('renders inline table content and vertical row spans', async () => {
    const table: TableNode<Record<string, unknown>> = {
      data: [{ _name_rowSpan: 2, name: 'Alpha' }, { name: 'Beta' }],
      type: 'table',
      columns: [
        {
          key: 'name',
          title: 'Name',
          render: value => [
            { text: `Cell: ${String(value)}`, type: 'text' },
            {
              children: [' details'],
              type: 'hyperlink',
              url: 'https://example.com/details',
            },
          ],
        },
      ],
    }
    const pkg = await renderPackage([table])
    const documentXml = await pkg.read('word/document.xml')

    expect(documentXml).toContain('Cell: Alpha')
    expect(documentXml).toContain(' details')
    expect(documentXml).not.toContain('[object Object]')
    expect(documentXml).toContain('<w:hyperlink')
    expect(documentXml).toContain('<w:vMerge w:val="restart"/>')
  })

  it('preserves rich list items and emits nested numbering levels', async () => {
    const pkg = await renderPackage([
      {
        type: 'numberedList',
        items: [
          { text: 'Parent' },
          {
            level: 1,
            children: [
              { text: 'Nested ', type: 'text' },
              {
                children: ['link'],
                type: 'hyperlink',
                url: 'https://example.com/nested',
              },
            ],
          },
        ],
      },
    ])

    const documentXml = await pkg.read('word/document.xml')
    const numberingXml = await pkg.read('word/numbering.xml')

    expect(documentXml).toContain('Parent')
    expect(documentXml).toContain('Nested ')
    expect(documentXml).toContain('link')
    expect(documentXml).toContain('<w:ilvl w:val="1"/>')
    expect(documentXml).toContain('<w:hyperlink')
    expect(numberingXml).toContain('<w:lvl w:ilvl="1"')
    expect(numberingXml).toContain('<w:lvlText w:val="%1.%2."/>')
  })
})
