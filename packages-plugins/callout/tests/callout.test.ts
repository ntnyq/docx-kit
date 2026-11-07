import assert from 'node:assert/strict'
import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Table } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { calloutPlugin } from '../src'

describe('calloutPlugin', () => {
  it('returns a plugin named "callout"', () => {
    expect(calloutPlugin().name).toBe('callout')
  })

  it('renders an info callout', () => {
    const result = calloutPlugin().render(
      { content: 'System update tonight.', type: 'info' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders a warning callout with title', () => {
    const result = calloutPlugin().render(
      { content: '请确认后再提交。', title: '注意', type: 'warning' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders all four types', () => {
    const types = ['danger', 'info', 'success', 'warning'] as const
    for (const type of types) {
      const result = calloutPlugin().render(
        { content: 'Test', type },
        createPluginTestContext(),
      )
      expect(result).toBeInstanceOf(Table)
    }
  })

  it('renders without title', () => {
    const result = calloutPlugin().render(
      { content: 'Done!', type: 'success' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('handles empty content', () => {
    const result = calloutPlugin().render(
      { content: '', type: 'info' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('exports padding inside the shaded border and keeps the callout together', async () => {
    const result = await calloutPlugin().render(
      { content: 'Body text', title: 'Note', type: 'info' },
      createPluginTestContext(),
    )
    assert.ok(result instanceof Table)
    const document = new Document({ sections: [{ children: [result] }] })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const xml = await archive.file('word/document.xml')?.async('string')

    expect(xml).toContain('<w:cantSplit/>')
    expect(xml).toContain('<w:tcMar>')
    expect(xml).toContain('<w:left w:type="dxa" w:w="180"/>')
    expect(xml).toContain('<w:top w:type="dxa" w:w="140"/>')
    expect(xml).toContain('w:fill="D6E4F0"')
    expect(xml).toContain('<w:left w:val="single" w:color="4472C4" w:sz="18"/>')
    expect(xml).toContain('Note')
    expect(xml).toContain('Body text')
  })
  it('resolves part tokens and precedence while isolating titles, content and calls', async () => {
    const defaults = {
      cell: { borderLeft: { width: 4 }, margin: '8pt 10pt' as const },
      content: { color: '#444444' as const },
      title: { bold: true, color: '#111111' as const, fontSize: 18 },
    }
    const snapshot = structuredClone(defaults)
    const plugin = calloutPlugin({ styles: defaults })
    const builder = createDocx({
      theme: { colors: { primary: '#ABCDEF' }, spacing: { md: 12 } },
    })
      .use(plugin)
      .plugin(
        'callout',
        {
          content: '$colors.primary',
          title: 'Custom title',
          type: 'info',
          styles: {
            title: { bold: false, color: '$colors.primary' },
            cell: {
              borderLeft: { color: '$colors.primary' },
              marginLeft: '$spacing.md',
            },
          },
        },
        { color: '#222222', fontSize: 14, marginBottom: 0 },
      )
      .plugin('callout', {
        content: 'Default body',
        title: 'Default title',
        type: 'info',
      })
    const zip = await JSZip.loadAsync(await builder.toBuffer())
    const xml = await zip.file('word/document.xml')?.async('string')
    const tables = xml?.match(/<w:tbl>.*?<\/w:tbl>/g) ?? []
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('w:val="single" w:color="ABCDEF" w:sz="32"')
    expect(tables[0]).toContain('<w:left w:type="dxa" w:w="240"/>')
    expect(tables[0]).toContain('<w:right w:type="dxa" w:w="200"/>')
    expect(tables[0]).toContain('<w:bottom w:type="dxa" w:w="0"/>')
    const runs = tables[0]?.match(/<w:r>.*?<\/w:r>/g) ?? []
    expect(runs[0]).toContain('<w:b w:val="false"/>')
    expect(runs[0]).toContain('<w:color w:val="ABCDEF"/>')
    expect(runs[0]).toContain('<w:sz w:val="28"/>')
    expect(runs[1]).toContain('<w:color w:val="222222"/>')
    expect(runs[1]).toContain('$colors.primary')
    expect(tables[1]).toContain('<w:color w:val="111111"/>')
    expect(tables[1]).toContain('<w:sz w:val="36"/>')
    expect(defaults).toEqual(snapshot)
  })
})
