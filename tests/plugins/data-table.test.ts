import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { dataTablePlugin } from '../../src/plugins/data-table'

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

const sampleData = [
  { age: 30, name: 'Alice', salary: 85000 },
  { age: 25, name: 'Bob', salary: 62000 },
]

describe('dataTablePlugin', () => {
  it('returns a plugin named "dataTable"', () => {
    expect(dataTablePlugin().name).toBe('dataTable')
  })

  it('renders a Table from object array', () => {
    const result = dataTablePlugin().render({ data: sampleData }, makeCtx())
    expect(result).toBeInstanceOf(Table)
  })

  it('renders a placeholder for empty data', () => {
    const result = dataTablePlugin().render({ data: [] }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with labels', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        labels: { age: '年龄', name: '姓名', salary: '薪资' },
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with format option', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        format: { salary: 'currency' },
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with striped rows', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        striped: true,
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with bordered = false', () => {
    const result = dataTablePlugin().render(
      {
        bordered: false,
        data: sampleData,
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with explicit alignment', () => {
    const result = dataTablePlugin().render(
      {
        align: { age: 'center', name: 'left', salary: 'right' },
        data: sampleData,
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('handles null and undefined values', () => {
    const result = dataTablePlugin().render(
      {
        data: [
          { a: 'ok', b: null, c: undefined },
          { a: 'also', b: 42, c: 'yes' },
        ],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('handles Chinese characters', () => {
    const result = dataTablePlugin().render(
      {
        data: [{ 状态: '进行中', 项目名称: 'XX系统' }],
        labels: { 状态: '状态', 项目名称: '项目名称' },
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })
})
