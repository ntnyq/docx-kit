import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph, Table } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { metricCardPlugin } from '../src'
import type { PluginRenderContext } from '@docxkit/core'
import type { MetricCardOptions } from '../src'

async function renderXml(
  options: MetricCardOptions,
  context?: Partial<PluginRenderContext>,
) {
  const result = await metricCardPlugin().render(
    options,
    createPluginTestContext(context),
  )
  const archive = await JSZip.loadAsync(
    await Packer.toBuffer(new Document({ sections: [{ children: [result] }] })),
  )
  return archive.file('word/document.xml')!.async('string')
}

describe('metricCardPlugin', () => {
  it('renders labels, zero values, units, comparisons, and trends', async () => {
    expect(metricCardPlugin().name).toBe('metricCard')
    const result = await metricCardPlugin().render(
      { items: [{ label: 'Revenue', value: 0 }] },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
    const xml = await renderXml({
      items: [
        {
          change: '+12%',
          label: 'Revenue',
          trend: 'up',
          unit: 'USD',
          value: 0,
        },
        { change: '-5%', label: 'Users', trend: 'down', value: '1,200' },
        {
          change: 'Unchanged',
          label: 'Uptime',
          trend: 'neutral',
          unit: '%',
          value: 99.9,
        },
      ],
    })
    for (const text of [
      'Revenue',
      '>0<',
      ' USD',
      '1,200',
      '99.9',
      '↑ +12%',
      '↓ -5%',
      '→ Unchanged',
    ]) {
      expect(xml).toContain(text)
    }
    expect(xml.match(/<w:tc>/g)).toHaveLength(3)
    expect(xml).toContain('<w:color w:val="2E7D32"/>')
    expect(xml).toContain('<w:color w:val="C62828"/>')
  })

  it('pads the final row, gives each cell equal width, and prevents row splits', async () => {
    const xml = await renderXml({
      columns: 2,
      items: [
        { label: 'A much longer metric label', value: '123,456,789,000' },
        { label: 'B', value: 2 },
        { label: 'C', value: 3 },
      ],
    })
    expect(xml.match(/<w:tr>/g)).toHaveLength(2)
    expect(xml.match(/<w:tc>/g)).toHaveLength(4)
    expect(xml.match(/<w:tcW w:type="pct" w:w="2500"\/>/g)).toHaveLength(4)
    expect(xml.match(/<w:cantSplit\/>/g)).toHaveLength(2)
    expect(xml).toContain('<w:tblLayout w:type="fixed"/>')
    expect(xml).toContain('<w:top w:val="none"')
    expect(xml.match(/w:fill="F4F6F8"/g)).toHaveLength(3)
  })

  it('caps columns to available items and supports a single column', async () => {
    for (const columns of [1, 100, Number.MAX_SAFE_INTEGER]) {
      const xml = await renderXml({
        columns,
        items: [{ label: 'Only metric', value: 42 }],
      })
      expect(xml.match(/<w:tc>/g)).toHaveLength(1)
      expect(xml).toContain('<w:tcW w:type="pct" w:w="5000"/>')
    }
  })

  it('shows comparisons without a trend and arrows without comparison text', async () => {
    const xml = await renderXml({
      items: [
        { change: 'Compared with last month', label: 'A', value: 1 },
        { label: 'B', trend: 'up', value: 2 },
      ],
    })
    expect(xml).toContain('Compared with last month')
    expect(xml).toContain('>↑<')
    expect(xml).not.toContain('undefined')
  })

  it('uses theme colors and lets invocation styles override all cards', async () => {
    const options: MetricCardOptions = {
      items: [
        { label: 'A', trend: 'up', value: 1 },
        { label: 'B', trend: 'down', value: 2 },
      ],
    }
    const config = {
      theme: {
        colors: {
          danger: '#AA0000',
          muted: '#789ABC',
          primary: '#123456',
          success: '#00AA00',
          surface: '#ABCDEF',
        },
      },
    }
    const themed = await renderXml(options, { config })
    for (const color of ['AA0000', '789ABC', '123456', '00AA00', 'ABCDEF']) {
      expect(themed).toContain(color)
    }
    const styled = await renderXml(options, {
      config,
      style: {
        backgroundColor: '#FEDCBA',
        color: '#123ABC',
        fontSize: 14,
        textAlign: 'right',
      },
    })
    expect(styled).toContain('<w:color w:val="123ABC"/>')
    expect(styled).toContain('<w:sz w:val="28"/>')
    expect(styled).toContain('w:fill="FEDCBA"')
    expect(styled).not.toContain('w:fill="ABCDEF"')
  })

  it('renders a placeholder for empty items', async () => {
    expect(
      await metricCardPlugin().render({ items: [] }, createPluginTestContext()),
    ).toBeInstanceOf(Paragraph)
    expect(await renderXml({ items: [] })).toContain('(no metrics)')
  })

  it.each([
    0,
    -1,
    1.5,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    Number.MAX_SAFE_INTEGER + 1,
  ])('rejects invalid columns %s even for empty items', async columns => {
    await expect(renderXml({ columns, items: [] })).rejects.toThrow(
      'positive safe integer',
    )
  })
})
