import assert from 'node:assert/strict'
import { createDocx } from '@docxkit/core'
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
  it('styles headers, body and stripes with per-level precedence', async () => {
    const plugin = dataTablePlugin({
      styles: {
        bodyCell: { backgroundColor: '#EEEEEE', margin: 8 },
        bodyText: { fontSize: 13 },
        headerText: { fontSize: 16 },
        stripedCell: { backgroundColor: '#DDDDDD', borderBottom: { width: 2 } },
      },
    })
    const builder = createDocx({ theme: { colors: { primary: '#123ABC' } } })
      .use(plugin)
      .plugin(
        'dataTable',
        {
          striped: true,
          data: [
            { label: 'First', score: 1 },
            { label: 'Second', score: 2 },
          ],
          styles: {
            bodyCell: { backgroundColor: '#CCCCCC', marginLeft: 0 },
            bodyParagraph: { textAlign: 'center' },
            headerCell: { backgroundColor: '#BBBBBB' },
            headerText: { bold: false, color: '$colors.primary' },
            stripedCell: {
              backgroundColor: '#AAAAAA',
              borderBottom: { color: '$colors.primary' },
            },
          },
        },
        { backgroundColor: '#FFFFFF', color: '#222222', fontSize: 12 },
      )
    const zip = await JSZip.loadAsync(await builder.toBuffer())
    const xml = await zip.file('word/document.xml')?.async('string')
    const rows = xml?.match(/<w:tr>.*?<\/w:tr>/g) ?? []
    expect(rows).toHaveLength(3)
    expect(rows[0]).toContain('w:fill="BBBBBB"')
    expect(rows[0]).toContain('<w:color w:val="123ABC"/>')
    expect(rows[0]).toContain('<w:b w:val="false"/>')
    expect(rows[0]).toContain('<w:sz w:val="24"/>')
    expect(rows[1]).toContain('w:fill="CCCCCC"')
    expect(rows[1]).not.toContain('w:fill="AAAAAA"')
    expect(rows[1]).toContain('<w:left w:type="dxa" w:w="0"/>')
    expect(rows[1]).toContain('<w:right w:type="dxa" w:w="160"/>')
    expect(rows[1]).toContain('<w:jc w:val="center"/>')
    expect(rows[2]).toContain('w:fill="AAAAAA"')
    expect(rows[2]).toContain('w:val="single" w:color="123ABC" w:sz="16"')
  })

  it('keeps stripe styles inactive when disabled and allows explicit borders', async () => {
    const builder = createDocx()
      .use(dataTablePlugin())
      .plugin('dataTable', {
        bordered: false,
        data: [{ a: 'First' }, { a: 'Second' }],
        styles: {
          stripedCell: { backgroundColor: '#ABCDEF' },
          bodyCell: {
            borderBottom: { color: '#123456', style: 'dashed', width: 1 },
          },
        },
      })
    const zip = await JSZip.loadAsync(await builder.toBuffer())
    const xml = await zip.file('word/document.xml')?.async('string')
    expect(xml).toContain('w:val="dashed" w:color="123456" w:sz="8"')
    expect(xml).not.toContain('w:fill="ABCDEF"')
  })

  it('styles the empty placeholder independently', async () => {
    const builder = createDocx()
      .use(
        dataTablePlugin({
          styles: {
            bodyText: { color: '#999999' },
            emptyText: { color: '#111111', italic: true },
          },
        }),
      )
      .plugin(
        'dataTable',
        {
          data: [],
          styles: {
            emptyParagraph: { textAlign: 'right' },
            emptyText: { color: '#333333' },
          },
        },
        { color: '#222222' },
      )
    const zip = await JSZip.loadAsync(await builder.toBuffer())
    const xml = await zip.file('word/document.xml')?.async('string')
    expect(xml).toContain('<w:color w:val="333333"/>')
    expect(xml).toContain('<w:i/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
  })
})
