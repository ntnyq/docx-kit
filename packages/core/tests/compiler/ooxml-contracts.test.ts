import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  TableNode,
} from '@docxkit/types'

async function renderPackage(
  nodes: BlockNode[],
  config: DocxKitConfig = {},
  plugins = new Map<string, DocxPlugin>(),
) {
  const document = await compileDocument({
    config,
    nodes,
    plugins,
  })
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document))

  async function read(path: string) {
    const file = archive.file(path)
    if (!file) {
      throw new Error(`Missing OOXML part: ${path}`)
    }
    return file.async('string')
  }

  async function readBytes(path: string) {
    const file = archive.file(path)
    if (!file) {
      throw new Error(`Missing OOXML part: ${path}`)
    }
    return file.async('uint8array')
  }

  return { read, readBytes }
}

describe('compiler OOXML contracts', () => {
  it('emits custom document properties and embedded fonts', async () => {
    const fontData = Uint8Array.from({ length: 64 }, (_, index) => index)
    const pkg = await renderPackage(
      [
        {
          children: [{ text: 'Embedded font', type: 'text' }],
          type: 'paragraph',
        },
      ],
      {
        fonts: [{ data: fontData, name: 'Docx Kit Test' }],
        metadata: {
          title: 'Property contract',
          customProperties: {
            Department: 'Engineering',
            Release: '2026.07',
          },
        },
      },
    )

    const customPropertiesXml = await pkg.read('docProps/custom.xml')
    const embeddedFont = await pkg.readBytes('word/fonts/font1.odttf')
    const fontTableXml = await pkg.read('word/fontTable.xml')
    const fontTableRelationships = await pkg.read(
      'word/_rels/fontTable.xml.rels',
    )

    expect(customPropertiesXml).toContain('name="Department"')
    expect(customPropertiesXml).toContain('<vt:lpwstr>Engineering</vt:lpwstr>')
    expect(customPropertiesXml).toContain('name="Release"')
    expect(fontTableXml).toContain('<w:font w:name="Docx Kit Test">')
    expect(fontTableXml).toContain('<w:embedRegular')
    expect(fontTableRelationships).toContain('Target="fonts/font1.odttf"')
    expect(embeddedFont).toHaveLength(fontData.length)
    expect(embeddedFont).not.toEqual(fontData)
  })

  it('emits advanced table layout, positioning, and cell styles', async () => {
    const table: TableNode<Record<string, unknown>> = {
      alignment: 'center',
      data: [{ amount: 42 }],
      layout: 'fixed',
      styleName: 'TableGrid',
      type: 'table',
      width: '60%',
      borders: {
        insideHorizontal: {
          color: '#94a3b8',
          style: 'dotted',
          width: '0.5pt',
        },
      },
      columns: [
        {
          align: 'right',
          key: 'amount',
          title: 'Amount',
          width: '40pt',
          cellStyle: {
            borderBottom: {
              color: '#123456',
              style: 'double',
              width: '1pt',
            },
          },
        },
      ],
      floating: {
        horizontalAnchor: 'page',
        overlap: false,
        verticalAnchor: 'margin',
        x: '1in',
        y: '2in',
      },
      tableLook: {
        firstRow: true,
        noHBand: false,
      },
      cellStyle: value => ({
        backgroundColor: Number(value) > 0 ? '#dcfce7' : '#fee2e2',
      }),
    }
    const pkg = await renderPackage([table])
    const documentXml = await pkg.read('word/document.xml')

    expect(documentXml).toContain('<w:tblStyle w:val="TableGrid"/>')
    expect(documentXml).toContain('<w:tblLayout w:type="fixed"/>')
    expect(documentXml).toContain('<w:jc w:val="center"/>')
    expect(documentXml).toContain('<w:tblW w:type="pct" w:w="60%"/>')
    expect(documentXml).toContain(
      '<w:tblpPr w:tblpX="1440" w:tblpY="2880" w:horzAnchor="page" w:vertAnchor="margin"><w:tblOverlap w:val="never"/></w:tblpPr>',
    )
    expect(documentXml).toContain(
      '<w:insideH w:val="dotted" w:color="94a3b8" w:sz="10"/>',
    )
    expect(documentXml).toContain(
      '<w:bottom w:val="double" w:color="123456" w:sz="20"/>',
    )
    expect(documentXml).toContain('<w:shd w:fill="dcfce7" w:val="clear"/>')
    expect(documentXml).toContain('<w:tcW w:type="dxa" w:w="800"/>')
  })

  it('emits footnote and comment document parts', async () => {
    const pkg = await renderPackage([
      {
        type: 'paragraph',
        children: [
          { text: 'Statement', type: 'text' },
          {
            content: ['Footnote body'],
            type: 'footnote',
          },
          {
            author: 'Ada Lovelace',
            children: [{ text: ' reviewed text', type: 'text' }],
            comment: ['Comment body'],
            date: '2026-07-26T00:00:00Z',
            initials: 'AL',
            type: 'comment',
          },
        ],
      },
    ])

    const commentsXml = await pkg.read('word/comments.xml')
    const documentXml = await pkg.read('word/document.xml')
    const footnotesXml = await pkg.read('word/footnotes.xml')

    expect(documentXml).toContain('<w:footnoteReference w:id="1"/>')
    expect(footnotesXml).toContain('<w:footnote w:id="1">')
    expect(footnotesXml).toContain('Footnote body')
    expect(documentXml).toContain('<w:commentRangeStart w:id="0"/>')
    expect(documentXml).toContain('<w:commentRangeEnd w:id="0"/>')
    expect(documentXml).toContain('<w:commentReference w:id="0"/>')
    expect(commentsXml).toContain('w:author="Ada Lovelace"')
    expect(commentsXml).toContain('w:initials="AL"')
    expect(commentsXml).toContain('Comment body')
  })

  it('emits image accessibility, floating, and aspect-ratio metadata', async () => {
    const image = new Uint8Array(24)
    image.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const view = new DataView(image.buffer)
    view.setUint32(16, 400)
    view.setUint32(20, 200)

    const pkg = await renderPackage([
      {
        alt: 'Quarterly revenue chart',
        data: image,
        floating: { wrap: 'square', x: 10, y: 20 },
        type: 'image',
        width: 100,
      },
    ])
    const documentXml = await pkg.read('word/document.xml')

    expect(documentXml).toContain('descr="Quarterly revenue chart"')
    expect(documentXml).toContain('<wp:anchor')
    expect(documentXml).toContain('cx="952500" cy="476250"')
    expect(documentXml).toContain('<wp:wrapSquare')
  })

  it('shares the compilation session with nodes rendered by plugins', async () => {
    const plugin: DocxPlugin = {
      name: 'nestedFootnote',
      render: async (_options, context) =>
        context.compileNode({
          content: ['Nested footnote body'],
          type: 'footnote',
        }),
    }
    const pkg = await renderPackage(
      [{ name: 'nestedFootnote', options: {}, type: 'plugin' }],
      {},
      new Map([[plugin.name, plugin]]),
    )

    const documentXml = await pkg.read('word/document.xml')
    const footnotesXml = await pkg.read('word/footnotes.xml')

    expect(documentXml).toContain('<w:footnoteReference w:id="1"/>')
    expect(footnotesXml).toContain('Nested footnote body')
  })

  it('emits semantic content nodes and tracked revisions', async () => {
    const pkg = await renderPackage([
      {
        children: ['Chapter one'],
        name: 'chapter_one',
        type: 'bookmark',
      },
      {
        anchor: 'chapter_one',
        children: ['Jump to chapter'],
        type: 'hyperlink',
      },
      {
        checked: true,
        label: 'Approved',
        type: 'checkbox',
      },
      {
        type: 'math',
        children: [
          {
            denominator: [{ text: '2', type: 'text' }],
            numerator: [{ text: '1', type: 'text' }],
            type: 'fraction',
          },
        ],
      },
      {
        author: 'Ada',
        children: ['new'],
        date: '2026-07-26T00:00:00Z',
        revisionId: 7,
        type: 'insertedText',
      },
      {
        author: 'Ada',
        children: ['old'],
        date: '2026-07-26T00:00:00Z',
        revisionId: 8,
        type: 'deletedText',
      },
      {
        text: 'Text box',
        type: 'textBox',
        box: {
          height: '40pt',
          position: 'absolute',
          width: '120pt',
        },
      },
      { type: 'thematicBreak' },
    ])

    const documentXml = await pkg.read('word/document.xml')
    const settingsXml = await pkg.read('word/settings.xml')

    expect(documentXml).toContain('w:name="chapter_one"')
    expect(documentXml).toContain('w:anchor="chapter_one"')
    expect(documentXml).toContain('<w14:checkbox>')
    expect(documentXml).toContain('<m:oMath>')
    expect(documentXml).toContain('<m:f>')
    expect(documentXml).toContain('<w:ins w:id="7"')
    expect(documentXml).toContain('<w:del w:id="8"')
    expect(documentXml).toContain('<w:txbxContent>')
    expect(documentXml).toContain('Text box')
    expect(documentXml).toContain('<w:pBdr>')
    expect(settingsXml).toContain('<w:trackRevisions/>')
  })

  it('emits advanced run and paragraph style properties', async () => {
    const pkg = await renderPackage([
      {
        type: 'paragraph',
        children: [
          {
            text: 'Styled text',
            type: 'text',
            style: {
              backgroundColor: '#abcdef',
              doubleStrike: true,
              emboss: true,
              imprint: true,
              letterSpacing: '1pt',
              rightToLeft: true,
              underline: 'double',
            },
          },
        ],
        style: {
          outlineLevel: 2,
          widowControl: false,
          tabStops: [
            {
              leader: 'dot',
              position: '2in',
              type: 'right',
            },
          ],
        },
      },
    ])

    const documentXml = await pkg.read('word/document.xml')

    expect(documentXml).toContain('<w:widowControl w:val="false"/>')
    expect(documentXml).toContain('<w:outlineLvl w:val="2"/>')
    expect(documentXml).toContain(
      '<w:tab w:val="right" w:pos="2880" w:leader="dot"/>',
    )
    expect(documentXml).toContain('<w:dstrike/>')
    expect(documentXml).toContain('<w:rtl/>')
    expect(documentXml).toContain('<w:emboss/>')
    expect(documentXml).toContain('<w:imprint/>')
    expect(documentXml).toContain('<w:spacing w:val="20"/>')
    expect(documentXml).toContain('<w:u w:val="double"/>')
    expect(documentXml).toContain('<w:shd w:fill="abcdef" w:val="clear"/>')
  })

  it('emits advanced first-section layout properties and column breaks', async () => {
    const pkg = await renderPackage([
      {
        type: 'sectionBreak',
        config: {
          type: 'continuous',
          columns: {
            count: 2,
            separator: true,
            spacing: '12pt',
          },
          lineNumbers: {
            countBy: 5,
            distance: '6pt',
            restart: 'newSection',
            start: 3,
          },
          page: {
            footerDistance: '10mm',
            gutter: '5mm',
            headerDistance: '8mm',
            margin: '20mm',
            borders: {
              display: 'allPages',
              offsetFrom: 'page',
              zOrder: 'front',
              top: {
                color: '#336699',
                style: 'double',
                width: '1pt',
              },
            },
            pageNumber: {
              format: 'upperRoman',
              start: 7,
            },
          },
        },
      },
      { text: 'First column', type: 'paragraph' },
      { type: 'columnBreak' },
      { text: 'Second column', type: 'paragraph' },
    ])

    const documentXml = await pkg.read('word/document.xml')

    expect(documentXml.match(/<w:sectPr/g)).toHaveLength(1)
    expect(documentXml).toContain('<w:br w:type="column"/>')
    expect(documentXml).toContain('w:header="454"')
    expect(documentXml).toContain('w:footer="567"')
    expect(documentXml).toContain('w:gutter="283"')
    expect(documentXml).toContain(
      '<w:pgNumType w:start="7" w:fmt="upperRoman"/>',
    )
    expect(documentXml).toContain(
      '<w:pgBorders w:display="allPages" w:offsetFrom="page" w:zOrder="front">',
    )
    expect(documentXml).toContain(
      '<w:top w:val="double" w:color="336699" w:sz="20"/>',
    )
    expect(documentXml).toContain(
      '<w:lnNumType w:countBy="5" w:start="3" w:restart="newSection" w:distance="120"/>',
    )
    expect(documentXml).toContain(
      '<w:cols w:space="240" w:num="2" w:sep="true"/>',
    )
    expect(documentXml).toContain('<w:type w:val="continuous"/>')
  })

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
