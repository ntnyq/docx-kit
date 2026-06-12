import { describe, expect, it } from 'vitest'
import { DocxBuilder } from '../../src/builder/DocxBuilder'
import type { BulletItem } from '../../src/dsl/nodes'

describe('DocxBuilder - bulletList', () => {
  it('creates a bullet list with string items', () => {
    const doc = new DocxBuilder()
    doc.bulletList(['Item 1', 'Item 2', 'Item 3'])

    const json = doc.toJSON()
    expect(json.content).toHaveLength(1)
    expect(json.content[0]).toMatchObject({
      items: ['Item 1', 'Item 2', 'Item 3'],
      type: 'bulletList',
    })
  })

  it('creates a bullet list with structured items', () => {
    const items: BulletItem[] = [
      { text: 'Rich item 1' },
      { style: { color: '#ff0000' }, text: 'Rich item 2' },
    ]
    const doc = new DocxBuilder()
    doc.bulletList(items)

    const json = doc.toJSON()
    expect(json.content).toHaveLength(1)
    expect((json.content[0] as any).items).toEqual(items)
  })

  it('supports custom bullet character', () => {
    const doc = new DocxBuilder()
    doc.bulletList(['A', 'B'], { bullet: '\u25CB' })

    const json = doc.toJSON()
    expect((json.content[0] as any).bullet).toBe('\u25CB')
  })
})

describe('DocxBuilder - numberedList', () => {
  it('creates a numbered list', () => {
    const doc = new DocxBuilder()
    doc.numberedList(['First', 'Second', 'Third'])

    const json = doc.toJSON()
    expect(json.content).toHaveLength(1)
    expect(json.content[0]).toMatchObject({
      items: ['First', 'Second', 'Third'],
      type: 'numberedList',
    })
  })

  it('supports numbering format options', () => {
    const doc = new DocxBuilder()
    doc.numberedList(['A', 'B'], {
      numberingFormat: 'upperRoman',
      start: 5,
    })

    const json = doc.toJSON()
    expect((json.content[0] as any).numberingFormat).toBe('upperRoman')
    expect((json.content[0] as any).start).toBe(5)
  })
})

describe('DocxBuilder - hyperlink', () => {
  it('creates a hyperlink', () => {
    const doc = new DocxBuilder()
    doc.hyperlink('https://example.com', 'Click here')

    const json = doc.toJSON()
    expect(json.content).toHaveLength(1)
    expect(json.content[0]).toMatchObject({
      children: ['Click here'],
      type: 'hyperlink',
      url: 'https://example.com',
    })
  })
})
