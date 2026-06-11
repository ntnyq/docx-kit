import { describe, expect, it } from 'vitest'
import { DocxBuilder } from '../../src/builder/DocxBuilder'

describe('DocxBuilder - section', () => {
  it('inserts a sectionBreak node when .section() is called', () => {
    const doc = new DocxBuilder()
    doc.p('Section 1').section().p('Section 2')

    const json = doc.toJSON()
    expect(json.content).toHaveLength(3)
    expect(json.content[0]).toMatchObject({
      text: 'Section 1',
      type: 'paragraph',
    })
    expect(json.content[1]).toMatchObject({
      config: undefined,
      type: 'sectionBreak',
    })
    expect(json.content[2]).toMatchObject({
      text: 'Section 2',
      type: 'paragraph',
    })
  })

  it('stores section config in the break node', () => {
    const doc = new DocxBuilder()
    doc
      .section({ page: { orientation: 'landscape', size: 'A3' } })
      .p('Landscape content')

    const json = doc.toJSON()
    expect(json.content).toHaveLength(2)
    const breakNode = json.content[0] as any
    expect(breakNode.type).toBe('sectionBreak')
    expect(breakNode.config).toEqual({
      page: { orientation: 'landscape', size: 'A3' },
    })
  })

  it('supports section with header config', () => {
    const doc = new DocxBuilder()
    doc.p('Page 1')
    doc.section({
      footer: { default: { children: ['Page 1'] } },
      header: { default: { children: ['Chapter 1', 'Confidential'] } },
    })
    doc.p('Content with header')

    const json = doc.toJSON()
    const breakNode = json.content[1] as any
    expect(breakNode.type).toBe('sectionBreak')
    expect(breakNode.config).toEqual({
      footer: { default: { children: ['Page 1'] } },
      header: { default: { children: ['Chapter 1', 'Confidential'] } },
    })
  })

  it('supports multiple sections with different configs', () => {
    const doc = new DocxBuilder()
    doc.p('Section 1')
    doc.section({ page: { size: 'Letter' } })
    doc.p('Section 2')
    doc.section({ page: { orientation: 'landscape', size: 'A3' } })
    doc.p('Section 3')

    const json = doc.toJSON()
    const types = json.content.map((n: any) => n.type)
    expect(types).toEqual([
      'paragraph',
      'sectionBreak',
      'paragraph',
      'sectionBreak',
      'paragraph',
    ])
  })

  it('add() can insert a sectionBreak node directly', () => {
    const doc = new DocxBuilder()
    doc.add({ type: 'sectionBreak' } as any)
    doc.p('Content')

    const json = doc.toJSON()
    expect(json.content[0]).toMatchObject({ type: 'sectionBreak' })
  })
})
