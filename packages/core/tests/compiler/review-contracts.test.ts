import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import {
  compileBorderRule,
  compileParagraphStyle,
  compileTextStyle,
} from '../../src/compiler/compileStyle'
import type { BlockNode, DocxKitConfig, HighlightColor } from '@docxkit/types'

async function renderPackage(nodes: BlockNode[], config: DocxKitConfig = {}) {
  const document = await compileDocument({ config, nodes, plugins: new Map() })
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
  return async (path = 'word/document.xml') => {
    const part = archive.file(path)
    if (!part) {
      throw new Error(`Missing part: ${path}`)
    }
    return part.async('string')
  }
}

describe('reviewed document contracts', () => {
  it.each([
    ['Georgia, serif', 'Georgia'],
    ['Inter, Arial, sans-serif', 'Inter'],
    ['Garamond, Georgia, serif', 'Garamond'],
    ['JetBrains Mono, monospace', 'JetBrains Mono'],
    ['  "Quoted, Family", serif', 'Quoted, Family'],
    ["'Times New Roman', serif", 'Times New Roman'],
  ])('emits one Word font family for %s', async (fontFamily, expected) => {
    const read = await renderPackage([
      { style: { fontFamily }, text: 'Text', type: 'paragraph' },
    ])
    expect(await read()).toContain(`w:ascii="${expected}"`)
  })

  it('preserves the grid for combined spans, including fully covered rows', async () => {
    const read = await renderPackage([
      {
        header: false,
        type: 'table',
        columns: [
          { key: 'a', title: 'A' },
          { key: 'b', title: 'B' },
          { key: 'c', title: 'C' },
        ],
        data: [
          {
            _a_colSpan: 2,
            _a_rowSpan: 2,
            _c_rowSpan: 2,
            a: 'Merged',
            b: 'Covered',
            c: 'Right',
          },
          { a: 'Covered', b: 'Covered', c: 'Covered' },
          { a: 'Left', b: 'Middle', c: 'Last' },
        ],
      },
    ])
    const xml = await read()
    const rows = xml.match(/<w:tr>.*?<\/w:tr>/g) ?? []
    expect(xml.match(/<w:gridCol /g)).toHaveLength(3)
    expect(rows.map(row => row.match(/<w:tc>/g)?.length)).toEqual([2, 2, 3])
    expect(rows[1].match(/w:vMerge w:val="continue"/g)).toHaveLength(2)
    expect(rows[1]).toContain('<w:gridSpan w:val="2"/>')
    expect(xml).not.toContain('Covered')
  })

  it('keeps all grid columns when every row has a horizontal span', async () => {
    const read = await renderPackage([
      {
        data: [{ a: 'Merged', b: 'Covered' }],
        type: 'table',
        columns: [
          { colSpan: 2, key: 'a', title: 'Group' },
          { key: 'b', title: 'Covered header' },
        ],
      },
    ])
    const xml = await read()
    expect(xml.match(/<w:gridCol /g)).toHaveLength(2)
    expect(xml.match(/<w:tc>/g)).toHaveLength(2)
    expect(xml).not.toContain('Covered')
  })

  it.each([0, -1, 1.5, Infinity, '2'])(
    'rejects invalid row span %s',
    async rowSpan => {
      await expect(
        renderPackage([
          {
            columns: [{ key: 'a', title: 'A' }],
            data: [{ _a_rowSpan: rowSpan, a: 'Value' }],
            type: 'table',
          },
        ]),
      ).rejects.toThrow('positive integers')
    },
  )

  it('rejects overlapping spans', async () => {
    await expect(
      renderPackage([
        {
          type: 'table',
          columns: [
            { key: 'a', title: 'A' },
            { key: 'b', title: 'B' },
          ],
          data: [
            { _b_rowSpan: 2, a: 'Left', b: 'Right' },
            { _a_colSpan: 2, a: 'Overlap' },
          ],
        },
      ]),
    ).rejects.toThrow('spans overlap')
  })

  it('rejects column spans outside the grid', async () => {
    await expect(
      renderPackage([
        {
          columns: [{ colSpan: 2, key: 'a', title: 'A' }],
          data: [],
          type: 'table',
        },
      ]),
    ).rejects.toThrow('exceeds the table grid')
  })

  it('inherits paragraph text styles in inline bookmarks', async () => {
    const read = await renderPackage([
      {
        style: { color: '#123456' },
        type: 'paragraph',
        children: [
          { children: ['Inherited'], name: 'target', type: 'bookmark' },
        ],
      },
    ])
    expect(await read()).toContain('<w:color w:val="123456"/>')
  })

  it('omits cells covered by row spans and preserves their grid positions', async () => {
    const read = await renderPackage([
      {
        type: 'table',
        columns: [
          { key: 'name', title: 'Name' },
          { key: 'score', title: 'Score' },
        ],
        data: [
          { _name_rowSpan: 3, name: 'Alpha', score: 1 },
          { name: 'Covered', score: 2 },
          { score: 3 },
          { name: 'Delta', score: 4 },
        ],
      },
    ])
    const xml = await read()
    const rows = xml.match(/<w:tr>.*?<\/w:tr>/g) ?? []
    expect(rows).toHaveLength(5)
    for (const row of rows) {
      expect(row.match(/<w:tc>/g)).toHaveLength(2)
    }
    expect(rows[2]).toMatch(/w:vMerge w:val="continue".*?<w:t[^>]*>2<\/w:t>/)
    expect(rows[3]).toMatch(/w:vMerge w:val="continue".*?<w:t[^>]*>3<\/w:t>/)
    expect(xml).not.toContain('Covered')
    expect(rows[4]).toContain('Delta')
  })

  it('applies cell paragraph styles while allowing column alignment overrides', async () => {
    const read = await renderPackage([
      {
        cellStyle: { lineHeight: 2, textAlign: 'right' },
        data: [{ name: 'Alpha', score: 1 }],
        header: false,
        type: 'table',
        columns: [
          { key: 'name', title: 'Name' },
          { align: 'center', key: 'score', title: 'Score' },
        ],
      },
    ])
    const xml = await read()
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(
      xml.match(/<w:spacing w:line="480" w:lineRule="auto"\/>/g),
    ).toHaveLength(2)
  })

  it('resolves bookmark classes and preserves child style overrides', async () => {
    const read = await renderPackage(
      [
        {
          className: 'red',
          name: 'target',
          type: 'bookmark',
          children: [
            'Inherited',
            { style: { color: '#0000FF' }, text: 'Override', type: 'text' },
          ],
        },
      ],
      { styles: { red: { color: '#FF0000' } } },
    )
    const xml = await read()
    expect(xml).toContain('<w:color w:val="FF0000"/>')
    expect(xml).toContain('<w:color w:val="0000FF"/>')
  })

  it('preserves nested math when a script has no operands', async () => {
    const read = await renderPackage([
      {
        type: 'math',
        children: [
          {
            type: 'script',
            children: [
              { text: 'BASE', type: 'text' },
              {
                denominator: [{ text: '2', type: 'text' }],
                numerator: [{ text: '1', type: 'text' }],
                type: 'fraction',
              },
            ],
          },
        ],
      },
    ])
    const xml = await read()
    expect(xml).toContain('<m:t>BASE</m:t>')
    expect(xml).toContain('<m:f>')
  })

  it('keeps tables as block children in headers and footers', async () => {
    const children: BlockNode[] = [
      {
        columns: [{ key: 'name', title: 'Name' }],
        data: [{ name: 'Alice' }],
        type: 'table',
      },
    ]
    const read = await renderPackage([
      {
        type: 'sectionBreak',
        config: {
          footer: { default: { children } },
          header: { default: { children } },
        },
      },
    ])
    for (const path of ['word/header1.xml', 'word/footer1.xml']) {
      const xml = await read(path)
      expect(xml).toContain('<w:tbl>')
      expect(xml).not.toContain('<w:p><w:tbl>')
    }
  })

  it('enables configured first and even variants', async () => {
    const read = await renderPackage([
      {
        type: 'sectionBreak',
        config: {
          footer: { even: { children: ['Even'] } },
          header: { first: { children: ['First'] } },
        },
      },
    ])
    expect(await read()).toContain('<w:titlePg/>')
    expect(await read('word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
  })

  it('retains zero margins and applies orientation to the default page size', async () => {
    const read = await renderPackage([], {
      page: { margin: 0, orientation: 'landscape' },
    })
    const xml = await read()
    expect(xml).toContain(
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>',
    )
    expect(xml).toContain(
      '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"',
    )
  })

  it('uses eighth-points for border widths without losing fractional precision', () => {
    expect(compileBorderRule({ width: '1pt' }).size).toBe(8)
    expect(compileBorderRule({ width: '0.5pt' }).size).toBe(4)
    expect(compileBorderRule({ width: '0.0625pt' }).size).toBe(1)
    expect(compileBorderRule({ width: '1px' }).size).toBe(6)
  })

  it('distinguishes absolute line heights from multipliers', () => {
    expect(compileParagraphStyle({ lineHeight: '12pt' }).spacing).toMatchObject(
      { line: 240, lineRule: 'exact' },
    )
    expect(compileParagraphStyle({ lineHeight: 1.5 }).spacing).toMatchObject({
      line: 360,
      lineRule: 'auto',
    })
  })

  it.each<HighlightColor>([
    'black',
    'blue',
    'cyan',
    'darkBlue',
    'darkCyan',
    'darkGray',
    'darkGreen',
    'darkMagenta',
    'darkRed',
    'darkYellow',
    'green',
    'lightGray',
    'magenta',
    'red',
    'white',
    'yellow',
    'none',
  ])('preserves highlight %s', color => {
    expect(compileTextStyle({ highlight: color }).highlight).toBe(color)
  })

  it('applies text defaults and preserves paragraph and run overrides', async () => {
    const read = await renderPackage(
      [
        { text: 'Default', type: 'paragraph' },
        {
          style: { color: '#222222' },
          type: 'paragraph',
          children: [
            { style: { fontSize: 10 }, text: 'Override', type: 'text' },
          ],
        },
        { level: 1, text: 'Heading', type: 'heading' },
      ],
      {
        defaults: {
          text: { color: '#111111', fontFamily: 'Courier New', fontSize: 22 },
        },
      },
    )
    const xml = await read()
    expect(xml.match(/w:ascii="Courier New"/g)).toHaveLength(3)
    expect(xml.match(/<w:sz w:val="44"\/>/g)).toHaveLength(2)
    expect(xml).toContain('<w:sz w:val="20"/>')
    expect(xml).toContain('<w:color w:val="222222"/>')
  })

  it('applies inline-only table styling, element defaults and header typography', async () => {
    const read = await renderPackage(
      [
        {
          columns: [{ key: 'name', title: 'Name' }],
          data: [{ name: 'Alice' }],
          headerCellStyle: { bold: true, color: '#ABCDEF' },
          style: { color: '#123456' },
          type: 'table',
        },
      ],
      {
        defaults: {
          cell: { backgroundColor: '#EEEEEE' },
          table: { fontSize: 18 },
          text: { fontFamily: 'Courier New' },
        },
      },
    )
    const xml = await read()
    expect(xml).toContain('<w:color w:val="123456"/>')
    expect(xml).toContain('<w:color w:val="ABCDEF"/>')
    expect(xml).toContain('<w:b/>')
    const cells = xml.match(/<w:tcPr>.*?<\/w:tcPr>/g)
    expect(cells).toHaveLength(2)
    expect(cells?.every(cell => cell.includes('w:fill="EEEEEE"'))).toBe(true)
    expect(xml.match(/<w:sz w:val="36"\/>/g)).toHaveLength(2)
    expect(xml.match(/w:ascii="Courier New"/g)).toHaveLength(2)
  })
})
