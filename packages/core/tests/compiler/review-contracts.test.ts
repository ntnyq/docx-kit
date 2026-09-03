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
