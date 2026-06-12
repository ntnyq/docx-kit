import { describe, expect, it } from 'vitest'
import { createDocx, renderDocx } from '../../src/builder/createDocx'
import { DocxBuilder } from '../../src/builder/DocxBuilder'

describe('createDocx', () => {
  it('returns a DocxBuilder instance', () => {
    const builder = createDocx()
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with page size', () => {
    const builder = createDocx({ page: { size: 'A4' } })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with styles', () => {
    const builder = createDocx({
      styles: {
        body: { fontSize: '12pt' },
        h1: { fontSize: '24pt', fontWeight: 'bold' },
      },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with metadata', () => {
    const builder = createDocx({
      metadata: { creator: 'docx-kit', title: 'Report' },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with defaults', () => {
    const builder = createDocx({
      defaults: { paragraph: { fontSize: '11pt' } },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with theme', () => {
    const builder = createDocx({
      theme: { colors: { primary: '#333' }, fonts: { body: 'Arial' } },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('builder from createDocx can chain methods', () => {
    const builder = createDocx({ page: { size: 'Letter' } })
    builder.h1('Test').p('Body')
    expect(builder.toJSON().content).toHaveLength(2)
  })

  it('creates builder with full config', () => {
    const builder = createDocx({
      defaults: { paragraph: { fontSize: '11pt' } },
      metadata: { creator: 'test', description: 'desc', title: 'Doc' },
      page: { margin: '20mm', orientation: 'portrait', size: 'A4' },
      styles: { title: { fontSize: '28pt', fontWeight: 'bold' } },
      theme: { colors: { primary: '#333' }, fonts: { body: 'Calibri' } },
    })
    builder.h1('Title', { className: 'title' })
    expect(builder.toJSON().content).toHaveLength(1)
  })
})

describe('renderDocx', () => {
  it('returns a DocxBuilder from JSON schema', () => {
    const builder = renderDocx({
      content: [
        { level: 1, text: 'Title', type: 'heading' },
        { text: 'Content', type: 'paragraph' },
        { type: 'pageBreak' },
      ],
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
    expect(builder.toJSON().content).toHaveLength(3)
  })

  it('separates page config from content', () => {
    const builder = renderDocx({
      content: [{ className: 'h1', level: 1, text: 'Hello', type: 'heading' }],
      page: { margin: '20mm', size: 'A4' },
      styles: {
        h1: { fontSize: '24pt' },
      },
    })
    expect(builder.toJSON().content).toHaveLength(1)
    expect(builder.toJSON().content![0]).toMatchObject({
      text: 'Hello',
      type: 'heading',
    })
  })

  it('supports table nodes in schema', () => {
    const builder = renderDocx({
      content: [
        {
          columns: [{ key: 'name', title: 'Name' }],
          data: [{ name: 'Alice' }],
          type: 'table',
        },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(1)
  })

  it('supports image nodes in schema', () => {
    const builder = renderDocx({
      content: [
        {
          data: new Uint8Array([1, 2, 3]),
          height: 100,
          type: 'image',
          width: 200,
        },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(1)
  })

  it('supports plugin nodes in schema', () => {
    const builder = renderDocx({
      content: [{ name: 'test', options: {}, type: 'plugin' }],
    })
    const node = builder.toJSON().content![0]
    expect(node).toMatchObject({ name: 'test', type: 'plugin' })
  })
})
