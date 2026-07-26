import { createPluginTestContext } from '@docxkit/pdk'
import { describe, expect, it, vi } from 'vitest'

import { barcodePlugin } from '../src'

const { toBuffer } = vi.hoisted(() => ({
  toBuffer: vi.fn(),
}))

vi.mock('bwip-js', () => ({
  default: { toBuffer },
  toBuffer,
}))

function createPng(width: number, height: number) {
  const png = new Uint8Array(24)
  png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const view = new DataView(png.buffer)
  view.setUint32(16, width)
  view.setUint32(20, height)
  return png
}

describe('barcodePlugin', () => {
  it('returns a plugin with name "barcode"', () => {
    expect(barcodePlugin().name).toBe('barcode')
  })

  it('renders a Code 128 barcode with defaults', async () => {
    toBuffer.mockResolvedValueOnce(createPng(600, 180))

    const result = await barcodePlugin().render(
      { text: 'DOCX-KIT-2026' },
      createPluginTestContext(),
    )

    expect(Array.isArray(result)).toBe(true)
    expect(result).toHaveLength(1)
    expect(toBuffer).toHaveBeenCalledWith({
      backgroundcolor: 'FFFFFF',
      barcolor: '000000',
      bcid: 'code128',
      height: 12,
      includetext: true,
      rotate: 'N',
      scale: 3,
      text: 'DOCX-KIT-2026',
      textxalign: 'center',
    })
  })

  it('forwards format, colors, sizing, and caption options', async () => {
    toBuffer.mockResolvedValueOnce(createPng(240, 100))

    const result = await barcodePlugin().render(
      {
        alignment: 'right',
        backgroundColor: '#FEF3C7',
        barColor: '#1F2937',
        barHeight: 18,
        caption: 'Inventory item',
        format: 'ean13',
        includeText: false,
        rotate: 'R',
        scale: 4,
        text: '5901234123457',
        textColor: '#DC2626',
        width: 200,
      },
      createPluginTestContext(),
    )

    expect(result).toHaveLength(2)
    expect(toBuffer).toHaveBeenCalledWith({
      backgroundcolor: 'FEF3C7',
      barcolor: '1F2937',
      bcid: 'ean13',
      height: 18,
      includetext: false,
      rotate: 'R',
      scale: 4,
      text: '5901234123457',
      textcolor: 'DC2626',
      textxalign: 'center',
    })
  })

  it('rejects non-PNG renderer output', async () => {
    toBuffer.mockResolvedValueOnce(new Uint8Array([1, 2, 3]))

    await expect(
      barcodePlugin().render({ text: 'invalid' }, createPluginTestContext()),
    ).rejects.toThrow('valid PNG')
  })
})
