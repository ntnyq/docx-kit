import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { propertyTablePlugin } from '../src'

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
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders an empty items array', () => {
    const result = propertyTablePlugin().render(
      { items: [] },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('respects keyBold = false', () => {
    const result = propertyTablePlugin().render(
      { items: [{ key: 'Name', value: 'Alice' }], keyBold: false },
      createPluginTestContext(),
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
      createPluginTestContext(),
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
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })
})
