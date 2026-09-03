import assert from 'node:assert/strict'
import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph, Table } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { dataTablePlugin } from '../src'

const sampleData = [
  { age: 30, name: 'Alice', salary: 85000 },
  { age: 25, name: 'Bob', salary: 62000 },
]

describe('dataTablePlugin', () => {
  it('returns a plugin named "dataTable"', () => {
    expect(dataTablePlugin().name).toBe('dataTable')
  })

  it('renders a Table from object array', () => {
    const result = dataTablePlugin().render(
      { data: sampleData },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders a placeholder for empty data', () => {
    const result = dataTablePlugin().render(
      { data: [] },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with labels', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        labels: { age: '年龄', name: '姓名', salary: '薪资' },
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with format option', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        format: { salary: 'currency' },
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with striped rows', () => {
    const result = dataTablePlugin().render(
      {
        data: sampleData,
        striped: true,
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with bordered = false', async () => {
    const result = await dataTablePlugin().render(
      {
        bordered: false,
        data: sampleData,
      },
      createPluginTestContext(),
    )
    assert.ok(result instanceof Table)
    const document = new Document({ sections: [{ children: [result] }] })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const xml = await archive.file('word/document.xml')?.async('string')
    const borders = xml?.match(/<w:tblBorders>.*?<\/w:tblBorders>/)?.[0]
    expect(borders).toBeDefined()
    expect(borders).not.toContain('w:val="single"')
    expect(borders?.match(/w:val="none"/g)).toHaveLength(6)
  })

  it('renders with explicit alignment', () => {
    const result = dataTablePlugin().render(
      {
        align: { age: 'center', name: 'left', salary: 'right' },
        data: sampleData,
      },
      createPluginTestContext(),
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
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('handles Chinese characters', () => {
    const result = dataTablePlugin().render(
      {
        data: [{ 状态: '进行中', 项目名称: 'XX系统' }],
        labels: { 状态: '状态', 项目名称: '项目名称' },
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('infers each column alignment only once', () => {
    let propertyReads = 0
    const data = Array.from(
      { length: 100 },
      (_, index) =>
        new Proxy(
          { a: index, b: index + 1, c: index + 2 },
          {
            get(target, property, receiver) {
              if (typeof property === 'string' && property in target) {
                propertyReads++
              }
              return Reflect.get(target, property, receiver)
            },
          },
        ),
    )

    const result = dataTablePlugin().render({ data }, createPluginTestContext())

    expect(result).toBeInstanceOf(Table)
    expect(propertyReads).toBeLessThan(1000)
  })
})
