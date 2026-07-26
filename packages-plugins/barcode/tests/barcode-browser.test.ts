import { createPluginTestContext } from '@docxkit/pdk'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { barcodePlugin } from '../src'

const { toCanvas } = vi.hoisted(() => ({
  toCanvas: vi.fn(),
}))

vi.mock('bwip-js', () => ({
  default: { toCanvas },
  toCanvas,
}))

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('barcodePlugin browser renderer', () => {
  it('uses the canvas renderer when a Node buffer renderer is unavailable', async () => {
    const canvas = {
      height: 100,
      toDataURL: vi.fn(() => 'data:image/png;base64,iVBORw0KGgo='),
      width: 240,
    }
    toCanvas.mockReturnValueOnce(canvas)
    vi.stubGlobal('document', {
      createElement: vi.fn(() => canvas),
    })

    const result = await barcodePlugin().render(
      { text: 'BROWSER-2026' },
      createPluginTestContext(),
    )

    expect(result).toHaveLength(1)
    expect(toCanvas).toHaveBeenCalledWith(
      canvas,
      expect.objectContaining({
        bcid: 'code128',
        text: 'BROWSER-2026',
      }),
    )
    expect(canvas.toDataURL).toHaveBeenCalledWith('image/png')
  })
})
