import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer } from 'docx'
import JSZip from 'jszip'
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

import { echartsPlugin } from '../src'

// Set up window to force browser rendering path
beforeAll(() => {
  ;(globalThis as any).window = {} as any
})

// Create DOM container mock
const mockContainer: { style: CSSStyleDeclaration; remove: () => void } = {
  remove: vi.fn(),
  // eslint-disable-next-line @typescript-eslint/consistent-type-assertions
  style: {} as CSSStyleDeclaration,
}
const mockChart = {
  dispose: vi.fn(),
  getDataURL: vi.fn().mockReturnValue('data:image/png;base64,iVBORw0KGgo='),
  setOption: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  mockChart.getDataURL.mockReturnValue('data:image/png;base64,iVBORw0KGgo=')
})

// Mock echarts
vi.mock('echarts', () => ({
  init: vi.fn().mockReturnValue(mockChart),
  default: {
    init: vi.fn().mockReturnValue(mockChart),
  },
}))

// Mock document for browser path
vi.stubGlobal('document', {
  createElement: vi.fn().mockReturnValue(mockContainer),
  body: {
    append: vi.fn(),
  },
})

describe('echartsPlugin', () => {
  it('applies invocation styles to the chart paragraph and caption', async () => {
    const output = await echartsPlugin().render(
      { caption: 'Styled chart', option: {} },
      createPluginTestContext({
        style: { color: '#123ABC', fontSize: 19, textAlign: 'right' },
      }),
    )
    if (!Array.isArray(output)) {
      throw new TypeError('Expected chart paragraphs')
    }
    const document = new Document({ sections: [{ children: output }] })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const xml = await archive.file('word/document.xml')?.async('string')
    expect(xml).toContain('<w:color w:val="123ABC"/>')
    expect(xml).toContain('<w:sz w:val="38"/>')
    expect(xml?.match(/<w:jc w:val="right"\/>/g)).toHaveLength(2)
  })
  it('returns a plugin with name "echarts"', () => {
    const plugin = echartsPlugin()
    expect(plugin.name).toBe('echarts')
  })

  it('renders a chart with default dimensions', async () => {
    const plugin = echartsPlugin()
    const result = await plugin.render(
      {
        option: {
          series: [{ data: [1, 2], type: 'bar' }],
          title: { text: 'Test' },
          xAxis: { data: ['A', 'B'], type: 'category' },
          yAxis: { type: 'value' },
        },
      },
      createPluginTestContext({
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: () => new Uint8Array([1, 2, 3]),
          },
        },
      }),
    )
    expect(result).toBeDefined()
    expect(Array.isArray(result)).toBe(true)
    expect(mockChart.setOption).toHaveBeenCalled()
  })

  it('renders a chart with caption', async () => {
    const plugin = echartsPlugin()
    const result = await plugin.render(
      {
        caption: 'Figure 1: Data',
        option: {
          series: [{ data: [1], type: 'line' }],
          xAxis: { data: ['X'] },
          yAxis: {},
        },
      },
      createPluginTestContext({
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: () => new Uint8Array([1, 2, 3]),
          },
        },
      }),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(2)
  })

  it('renders with custom width/height', async () => {
    const plugin = echartsPlugin()
    const result = await plugin.render(
      {
        height: 600,
        imageType: 'png',
        renderer: 'canvas',
        width: 800,
        option: {
          series: [{ data: [1], type: 'line' }],
          xAxis: { data: ['X'] },
          yAxis: {},
        },
      },
      createPluginTestContext({
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: () => new Uint8Array([1, 2, 3]),
          },
        },
      }),
    )
    expect(result).toBeDefined()
  })

  it('renders SVG with the required PNG fallback', async () => {
    const plugin = echartsPlugin()
    mockChart.getDataURL
      .mockReturnValueOnce('data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=')
      .mockReturnValueOnce('data:image/png;base64,iVBORw0KGgo=')

    const result = await plugin.render(
      {
        imageType: 'svg',
        option: {
          series: [{ data: [1], type: 'line' }],
          xAxis: { data: ['X'] },
          yAxis: {},
        },
      },
      createPluginTestContext({
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: () => new Uint8Array([1, 2, 3]),
          },
        },
      }),
    )

    expect(result).toBeDefined()
    expect(mockChart.getDataURL).toHaveBeenCalledWith(
      expect.objectContaining({ type: 'svg' }),
    )
    expect(mockChart.getDataURL).toHaveBeenCalledWith(
      expect.objectContaining({ type: 'png' }),
    )
  })

  it('rejects mismatched renderer and image formats', async () => {
    await expect(
      echartsPlugin().render(
        {
          imageType: 'svg',
          option: {},
          renderer: 'canvas',
        },
        createPluginTestContext(),
      ),
    ).rejects.toThrow('requires renderer "svg"')
  })

  it.each(['setOption', 'getDataURL'] as const)(
    'disposes the chart and container when %s fails',
    async method => {
      const failure = new Error('Chart failed')
      mockChart[method].mockImplementationOnce(() => {
        throw failure
      })
      await expect(
        echartsPlugin().render({ option: {} }, createPluginTestContext()),
      ).rejects.toBe(failure)
      expect(mockChart.dispose).toHaveBeenCalledExactlyOnceWith()
      expect(mockContainer.remove).toHaveBeenCalledExactlyOnceWith()
    },
  )

  it('removes the container when chart initialization fails', async () => {
    const { init } = await import('echarts')
    const failure = new Error('Initialization failed')
    vi.mocked(init).mockImplementationOnce(() => {
      throw failure
    })
    await expect(
      echartsPlugin().render({ option: {} }, createPluginTestContext()),
    ).rejects.toBe(failure)
    expect(mockChart.dispose).not.toHaveBeenCalled()
    expect(mockContainer.remove).toHaveBeenCalledExactlyOnceWith()
  })

  it('cleans up both instances when the SVG raster fallback fails', async () => {
    const failure = new Error('Fallback failed')
    mockChart.getDataURL.mockReturnValueOnce(
      'data:image/svg+xml;base64,PHN2Zy8+',
    )
    mockChart.setOption
      .mockImplementationOnce(() => {})
      .mockImplementationOnce(() => {
        throw failure
      })
    await expect(
      echartsPlugin().render(
        { imageType: 'svg', option: {} },
        createPluginTestContext(),
      ),
    ).rejects.toBe(failure)
    expect(mockChart.dispose).toHaveBeenCalledTimes(2)
    expect(mockContainer.remove).toHaveBeenCalledExactlyOnceWith()
  })
})
