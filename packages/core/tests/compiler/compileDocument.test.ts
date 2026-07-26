import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'

describe('compileDocument', () => {
  it('compiles empty document', async () => {
    const doc = await compileDocument({
      config: {},
      nodes: [],
      plugins: new Map(),
    })
    expect(doc).toBeDefined()
    expect(typeof (doc as any).constructor).toBe('function')
  })

  it('compiles document with multiple node types', async () => {
    const doc = await compileDocument({
      plugins: new Map(),
      config: {
        metadata: { creator: 'docx-kit', title: 'Test' },
        page: { margin: '20mm', size: 'A4' },
      },
      nodes: [
        { level: 1, text: 'Title', type: 'heading' },
        { text: 'Body text', type: 'paragraph' },
        { type: 'pageBreak' },
        { level: 2, text: 'Section 2', type: 'heading' },
      ],
    })
    expect(doc).toBeDefined()
  })

  it('compiles document with table nodes', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'table',
          columns: [
            { key: 'id', title: 'ID' },
            { key: 'name', title: 'Name' },
          ],
          data: [
            { id: 1, name: 'Alice' },
            { id: 2, name: 'Bob' },
          ],
        },
      ],
    })
    expect(doc).toBeDefined()
  })

  it('compiles document with image node', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          data: new Uint8Array([1, 2, 3]),
          height: 200,
          imageType: 'png',
          type: 'image',
          width: 300,
        },
      ],
    })
    expect(doc).toBeDefined()
  })

  it('compiles document with plugin node (registered)', async () => {
    const plugin = { name: 'test', render: () => 'content' }
    const plugins = new Map([['test', plugin]])
    const doc = await compileDocument({
      config: {},
      nodes: [{ name: 'test', options: {}, type: 'plugin' }],
      plugins,
    })
    expect(doc).toBeDefined()
  })

  it('handles array return from compileNode (plugin returning multiple items)', async () => {
    const plugin = {
      name: 'multi',
      render: () => [{ type: 'paragraph' }],
    }
    const doc = await compileDocument({
      config: {},
      nodes: [{ name: 'multi', options: {}, type: 'plugin' }],
      plugins: new Map([['multi', plugin]]),
    })
    expect(doc).toBeDefined()
  })

  it('sets page size to A4 landscape', async () => {
    const doc = await compileDocument({
      config: { page: { orientation: 'landscape', size: 'A4' } },
      nodes: [{ text: 'Landscape text', type: 'paragraph' }],
      plugins: new Map(),
    })
    expect(doc).toBeDefined()
  })

  it('sets page size to custom dimensions', async () => {
    const doc = await compileDocument({
      nodes: [{ text: 'Custom page', type: 'paragraph' }],
      plugins: new Map(),
      config: {
        page: {
          margin: '10pt',
          size: { height: '15cm', width: '10cm' },
        },
      },
    })
    expect(doc).toBeDefined()
  })

  it('sets page size to Letter', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'Letter' } },
      nodes: [],
      plugins: new Map(),
    })
    expect(doc).toBeDefined()
  })

  it('sets page size to A3', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'A3' } },
      nodes: [],
      plugins: new Map(),
    })
    expect(doc).toBeDefined()
  })

  it('sets page size to Legal', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'Legal' } },
      nodes: [],
      plugins: new Map(),
    })
    expect(doc).toBeDefined()
  })
})
