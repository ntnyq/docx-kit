import assert from 'node:assert/strict'
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
})
