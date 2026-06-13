/**
 * ECharts plugin — embeds interactive ECharts charts as static images.
 *
 * Renders charts in the browser via the DOM. For Node.js rendering,
 * a custom canvas/SVG implementation is needed (not built-in).
 *
 * @module plugins/echarts
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(echartsPlugin())
 *   .h1('Sales Chart')
 *   .plugin('echarts', {
 *     option: {
 *       title: { text: 'Monthly Sales' },
 *       xAxis: { data: ['Jan', 'Feb', 'Mar'] },
 *       yAxis: {},
 *       series: [{ type: 'bar', data: [120, 200, 150] }],
 *     },
 *     width: 640,
 *     height: 360,
 *     caption: 'Figure 1: Monthly sales data',
 *   })
 *   .save('chart.docx')
 * ```
 */

import {
  createImageRun,
  dataUrlToUint8Array,
  definePlugin,
  DocxKitError,
  ERROR_CODES,
} from '@docxkit/core'
import { Paragraph } from 'docx'

// Use type-only import so the dep stays optional at runtime
import type { EChartsOption } from 'echarts'

/**
 * Options for the ECharts plugin.
 */
export interface EChartsPluginOptions {
  /** Full ECharts option object (series, axes, title, etc.). */
  option: EChartsOption
  /** Optional caption text displayed below the chart. */
  caption?: string
  /** Chart height in pixels. @default 360 */
  height?: number
  /** Output image format. @default "png" */
  imageType?: 'png' | 'svg'
  /** ECharts rendering engine. @default "canvas" */
  renderer?: 'canvas' | 'svg'
  /** Chart width in pixels. @default 640 */
  width?: number
}

/** Internal render configuration (resolved defaults). */
interface RenderConfig {
  height: number
  imageType: 'png' | 'svg'
  renderer: 'canvas' | 'svg'
  width: number
}

/** Result of rendering a chart to an image. */
interface RenderImageResult {
  /** Raw image bytes. */
  data: Uint8Array
  /** Image format (`"png"` or `"svg"`). */
  type: 'png' | 'svg'
}

/**
 * Create an ECharts plugin instance.
 *
 * The plugin renders charts using the browser DOM. In Node.js
 * environments, it throws an error prompting the user to provide
 * a server-side canvas implementation.
 *
 * @returns A configured DocxPlugin for `'echarts'`
 *
 * @example
 * ```ts
 * import { createDocx, echartsPlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(echartsPlugin())
 *   .plugin('echarts', {
 *     option: {
 *       title: { text: 'Revenue' },
 *       xAxis: { type: 'category', data: ['Q1', 'Q2', 'Q3', 'Q4'] },
 *       yAxis: { type: 'value' },
 *       series: [{ data: [820, 932, 901, 1347], type: 'line' }],
 *     },
 *     caption: 'Quarterly revenue trend',
 *   })
 * ```
 */
export function echartsPlugin() {
  return definePlugin<'echarts', EChartsPluginOptions>({
    name: 'echarts',
    async render(options) {
      const width = options.width ?? 640
      const height = options.height ?? 360
      const imageType = options.imageType ?? 'png'
      const renderer = options.renderer ?? 'canvas'

      const image = await renderEChartsToImage(options.option, {
        height,
        imageType,
        renderer,
        width,
      })

      const imageRun = createImageRun({
        data: image.data,
        height,
        type: image.type,
        width,
      })

      const paragraphs: Paragraph[] = [new Paragraph({ children: [imageRun] })]

      if (options.caption) {
        paragraphs.push(new Paragraph({ text: options.caption }))
      }

      return paragraphs
    },
  })
}

// ---------- Internal helpers ----------

/**
 * Render an ECharts option to an image.
 *
 * Automatically selects browser or Node.js rendering based on
 * whether the `window` global is available.
 *
 * @param option - — Full ECharts option object
 * @param config - — Render dimensions and format
 * @returns Rendered image data and type
 */
async function renderEChartsToImage(
  option: EChartsOption,
  config: RenderConfig,
): Promise<RenderImageResult> {
  if (typeof window !== 'undefined') {
    return renderInBrowser(option, config)
  }
  return renderInNode(option, config)
}

/**
 * Browser-side ECharts rendering.
 *
 * Creates a hidden DOM container, renders the chart via `echarts.init()`,
 * extracts the image as a base64 data-URL, then cleans up the DOM.
 *
 * @param option - — Full ECharts option object
 * @param config - — Render dimensions and format
 * @returns Rendered image data and type
 */
async function renderInBrowser(
  option: EChartsOption,
  config: RenderConfig,
): Promise<RenderImageResult> {
  const echarts = await import('echarts')

  const container = document.createElement('div')
  container.style.cssText = `width:${config.width}px;height:${config.height}px;position:fixed;left:-99999px;`
  document.body.append(container)

  const chart = echarts.init(container, undefined, {
    height: config.height,
    renderer: config.renderer,
    width: config.width,
  })
  chart.setOption({ ...option, animation: false })

  const dataUrl = chart.getDataURL({
    backgroundColor: '#ffffff',
    pixelRatio: 2,
    type: config.imageType === 'svg' ? 'svg' : 'png',
  })

  chart.dispose()
  container.remove()

  return { data: await dataUrlToUint8Array(dataUrl), type: config.imageType }
}

/**
 * Node.js-side ECharts rendering (not built-in).
 *
 * Requires a server-side canvas or SVG DOM implementation
 * (e.g. `node-canvas` + `echarts`). Throws a PLUGIN_RENDER_FAILED
 * error prompting the user to provide a custom render function.
 *
 * @param _option - — ECharts option (unused — throws instead)
 * @param _config - — Render config (unused — throws instead)
 * @throws {DocxKitError} Always — Node rendering is not built in
 */
async function renderInNode(
  _option: EChartsOption,
  _config: RenderConfig,
): Promise<RenderImageResult> {
  throw new DocxKitError(
    ERROR_CODES.PLUGIN_RENDER_FAILED,
    'ECharts Node renderer is not built-in. '
      + 'Use a server-side canvas implementation (e.g. node-canvas + echarts) '
      + 'and provide a custom render function. '
      + 'See https://github.com/ntnyq/docx-kit for alternatives.',
  )
}
