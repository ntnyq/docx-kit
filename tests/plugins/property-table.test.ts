import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { propertyTablePlugin } from '../../src/plugins/property-table'

const makeCtx = () =>
  ({
    config: {},
    utils: {
      image: {
        fromBlob: async () => new Uint8Array(),
        fromDataUrl: async () => new Uint8Array(),
      },
    },
    compileNode: async () => null,
  }) as any

describe('propertyTablePlugin', () => {
  it('returns a plugin named "propertyTable"', () => {
    expect(propertyTablePlugin().name).toBe('propertyTable')
  })

  it('renders a Table with key-value pairs', () => {
    const result = propertyTablePlugin().render(
      {
        items: [
          { key: 'Name', value: 'Alice' },
          { key: 'Age', value: '30' },
        ],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders an empty items array', () => {
    const result = propertyTablePlugin().render({ items: [] }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('respects keyBold = false', () => {
    const result = propertyTablePlugin().render(
      { items: [{ key: 'Name', value: 'Alice' }], keyBold: false },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('respects striped = false', () => {
    const result = propertyTablePlugin().render(
      {
        striped: false,
        items: [
          { key: 'A', value: '1' },
          { key: 'B', value: '2' },
        ],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('handles special characters in keys and values', () => {
    const result = propertyTablePlugin().render(
      {
        items: [
          { key: '项目名称', value: 'XX管理系统' },
          { key: '技术栈', value: 'React & Node.js 5.0' },
        ],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })
})
