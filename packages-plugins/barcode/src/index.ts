/**
 * Barcode plugin — embeds a linear barcode image in the document.
 *
 * Uses the optional `bwip-js` peer dependency and selects its Node.js or
 * browser PNG renderer at runtime.
 *
 * @module plugins/barcode
 */

import {
  createImageRun,
  dataUrlToUint8Array,
  definePlugin,
} from '@docxkit/core'
import { AlignmentType, Paragraph } from 'docx'
import type { LiteralUnion } from '@docxkit/core'

/** Common linear barcode formats, while retaining the full bwip-js catalog. */
export type BarcodeFormat = LiteralUnion<
  | 'code128'
  | 'code39'
  | 'code93'
  | 'ean13'
  | 'ean8'
  | 'interleaved2of5'
  | 'isbn'
  | 'itf14'
  | 'upca'
  | 'upce'
>

/** Options for the barcode plugin. */
export interface BarcodeOptions {
  /** Value encoded by the barcode. */
  text: string
  /** Paragraph alignment. @default "center" */
  alignment?: 'center' | 'left' | 'right'
  /** Background color as a hex string. @default "#FFFFFF" */
  backgroundColor?: string
  /** Bar color as a hex string. @default "#000000" */
  barColor?: string
  /** Bar height in millimeters. @default 12 */
  barHeight?: number
  /** Optional caption below the barcode. */
  caption?: string
  /** Barcode symbology. @default "code128" */
  format?: BarcodeFormat
  /** Include the human-readable value below the bars. @default true */
  includeText?: boolean
  /** Barcode rotation. @default "N" */
  rotate?: 'I' | 'L' | 'N' | 'R'
  /** Rasterization scale. @default 3 */
  scale?: number
  /** Human-readable text color as a hex string. */
  textColor?: string
  /** Display width in pixels. Defaults to the natural width, capped at 320. */
  width?: number
}

interface BarcodeRenderer {
  toBuffer?: (options: BarcodeRenderOptions) => Promise<Uint8Array>
  toCanvas?: (
    canvas: HTMLCanvasElement,
    options: BarcodeRenderOptions,
  ) => HTMLCanvasElement
}

interface BarcodeRenderOptions {
  backgroundcolor: string
  barcolor: string
  bcid: string
  height: number
  includetext: boolean
  rotate: 'I' | 'L' | 'N' | 'R'
  scale: number
  text: string
  textxalign: 'center'
  textcolor?: string
}

interface RenderedBarcode {
  data: Uint8Array
  height: number
  width: number
}

/**
 * Create a barcode plugin instance.
 *
 * @returns A configured plugin for the `barcode` node name
 */
export function barcodePlugin() {
  return definePlugin<'barcode', BarcodeOptions>({
    name: 'barcode',
    async render(options) {
      const rendered = await renderBarcode(options)
      const width = options.width ?? Math.min(rendered.width, 320)
      const height = Math.max(
        1,
        Math.round(rendered.height * (width / rendered.width)),
      )
      const alignment = toAlignment(options.alignment)
      const paragraphs = [
        new Paragraph({
          alignment,
          children: [
            createImageRun({
              data: rendered.data,
              height,
              type: 'png',
              width,
            }),
          ],
        }),
      ]

      if (options.caption) {
        paragraphs.push(
          new Paragraph({
            alignment,
            text: options.caption,
          }),
        )
      }

      return paragraphs
    },
  })
}

async function loadBarcodeRenderer(): Promise<BarcodeRenderer> {
  try {
    const module = await import('bwip-js')
    return (module.default ?? module) as BarcodeRenderer
  } catch (err) {
    throw new Error(
      'The barcode plugin requires the optional "bwip-js" package. Install it before rendering barcodes.',
      { cause: err },
    )
  }
}

function normalizeColor(color: string): string {
  return color.replace(/^#/, '')
}

function readPngDimensions(data: Uint8Array) {
  const isPng =
    data.length >= 24
    && data[0] === 0x89
    && data[1] === 0x50
    && data[2] === 0x4e
    && data[3] === 0x47

  if (!isPng) {
    throw new Error('bwip-js did not return a valid PNG barcode image.')
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  return {
    height: view.getUint32(20),
    width: view.getUint32(16),
  }
}

async function renderBarcode(
  options: BarcodeOptions,
): Promise<RenderedBarcode> {
  const renderer = await loadBarcodeRenderer()
  const renderOptions: BarcodeRenderOptions = {
    backgroundcolor: normalizeColor(options.backgroundColor ?? '#FFFFFF'),
    barcolor: normalizeColor(options.barColor ?? '#000000'),
    bcid: options.format ?? 'code128',
    height: options.barHeight ?? 12,
    includetext: options.includeText ?? true,
    rotate: options.rotate ?? 'N',
    scale: options.scale ?? 3,
    text: options.text,
    textxalign: 'center',
    ...(options.textColor
      ? { textcolor: normalizeColor(options.textColor) }
      : {}),
  }

  if (renderer.toBuffer) {
    const data = new Uint8Array(await renderer.toBuffer(renderOptions))
    return { data, ...readPngDimensions(data) }
  }

  if (renderer.toCanvas && typeof document !== 'undefined') {
    const canvas = renderer.toCanvas(
      document.createElement('canvas'),
      renderOptions,
    )
    const data = await dataUrlToUint8Array(canvas.toDataURL('image/png'))
    return {
      data,
      height: canvas.height,
      width: canvas.width,
    }
  }

  throw new Error('No compatible bwip-js PNG renderer is available.')
}

function toAlignment(alignment: BarcodeOptions['alignment']) {
  switch (alignment) {
    case 'left':
      return AlignmentType.LEFT
    case 'right':
      return AlignmentType.RIGHT
    default:
      return AlignmentType.CENTER
  }
}
