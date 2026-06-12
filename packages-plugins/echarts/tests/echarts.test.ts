import { createPluginTestContext } from '@docxkit/pdk'
import { beforeAll, describe, expect, it, vi } from 'vitest'

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

  it('renders with svg type (not directly supported by ImageRun)', async () => {
    // SVG is not a standard ImageRun image type. The plugin attempts to
    // pass it through, but ImageRun will reject it at runtime. This is
    // expected behavior — customers should use png for charts.
    const plugin = echartsPlugin()
    mockChart.getDataURL.mockReturnValue(
      'data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=',
    )
    await expect(
      plugin.render(
        {
          imageType: 'svg' as any,
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
      ),
    ).rejects.toThrow()
  })
})
