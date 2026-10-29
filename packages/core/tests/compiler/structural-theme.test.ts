import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import type { BlockNode, DocxKitConfig, UnitValue } from '@docxkit/types'

const imageData =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

async function renderXml(nodes: BlockNode[], config: DocxKitConfig) {
  const document = await compileDocument({ config, nodes, plugins: new Map() })
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
  return archive.file('word/document.xml')!.async('string')
}

describe('structural theme tokens', () => {
  it('resolves page and section lengths and borders without changing content', async () => {
    const xml = await renderXml(
      [
        {
          type: 'sectionBreak',
          config: {
            columns: { columns: [{ width: '$spacing.column' }] },
            lineNumbers: { distance: '$spacing.distance' },
            page: { gutter: '$spacing.distance' },
          },
        },
        { text: '$colors.primary', type: 'paragraph' },
      ],
      {
        page: {
          headerDistance: '$spacing.distance',
          margin: '$spacing.margin',
          size: { height: '$spacing.height', width: '$spacing.width' },
          borders: {
            top: { color: '$colors.primary', width: '$spacing.border' },
          },
        },
        theme: {
          colors: { primary: '#123456' },
          spacing: {
            border: 1,
            column: 100,
            distance: 12,
            height: 600,
            margin: 20,
            width: 400,
          },
        },
      },
    )
    expect(xml).toContain('w:w="8000" w:h="12000"')
    expect(xml).toContain(
      'w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="240"',
    )
    expect(xml).toContain('w:color="123456" w:sz="8"')
    expect(xml).toContain('<w:col w:w="2000"')
    expect(xml).toContain('w:distance="240"')
    expect(xml).toContain('>$colors.primary</w:t>')
  })

  it('resolves block and inline image dimensions, floating positions, tables and text boxes', async () => {
    const xml = await renderXml(
      [
        {
          data: imageData,
          floating: { x: '$spacing.offset' },
          type: 'image',
          width: '$spacing.width',
        },
        {
          type: 'paragraph',
          children: [
            { data: imageData, type: 'image', width: '$spacing.width' },
          ],
        },
        {
          box: { height: '$spacing.offset', width: '$spacing.width' },
          text: 'Aside',
          type: 'textBox',
        },
        {
          columns: [{ key: 'value', title: 'Value', width: '$spacing.width' }],
          data: [{ value: '$colors.primary' }],
          floating: { x: '$spacing.offset' },
          type: 'table',
          width: '$spacing.width',
          borders: {
            top: { color: '$colors.primary', width: '$spacing.border' },
          },
        },
      ],
      {
        theme: {
          colors: { primary: '#123456' },
          spacing: { border: 1, offset: 12, width: 100 },
        },
      },
    )
    expect(xml.match(/<wp:extent cx="952500" cy="952500"\/>/g)).toHaveLength(2)
    expect(xml).toContain('<wp:posOffset>114300</wp:posOffset>')
    expect(xml).toContain('<w:tblW w:type="dxa" w:w="2000"/>')
    expect(xml).toContain('<w:tcW w:type="dxa" w:w="2000"/>')
    expect(xml).toContain('w:color="123456" w:sz="8"')
    expect(xml).toContain('>$colors.primary</w:t>')
    expect(xml).not.toContain('NaN')
  })

  it.each<UnitValue>([NaN, Infinity, '$spacing.missing'])(
    'rejects non-finite or unresolved image dimensions: %s',
    async width => {
      await expect(
        renderXml([{ data: imageData, type: 'image', width }], {}),
      ).rejects.toThrow(/length/i)
    },
  )
})
