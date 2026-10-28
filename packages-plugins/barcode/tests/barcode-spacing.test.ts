import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { expect, it } from 'vitest'
import { barcodePlugin } from '../src'

it('leaves a visible gap between Code 128 bars and the embedded text', async () => {
  const result = await barcodePlugin().render(
    { text: 'DOCX-KIT-2026' },
    createPluginTestContext(),
  )
  if (
    !Array.isArray(result)
    || !result.every(child => child instanceof Paragraph)
  ) {
    throw new TypeError('Expected barcode paragraphs')
  }
  const doc = new Document({ sections: [{ children: result }] })
  const zip = await JSZip.loadAsync(await Packer.toBuffer(doc))
  const image = Object.values(zip.files).find(file =>
    file.name.endsWith('.png'),
  )
  if (!image) {
    throw new Error('Missing barcode PNG')
  }
  const png = PNG.sync.read(await image.async('nodebuffer'))
  const inkRows: number[] = []
  for (let y = 0; y < png.height; y++) {
    for (let x = 0; x < png.width; x++) {
      const offset = (y * png.width + x) * 4
      if (png.data[offset + 3] > 0 && png.data[offset] < 128) {
        inkRows.push(y)
        break
      }
    }
  }
  const gapEnd = inkRows.findIndex(
    (row, index) => index > 0 && row > inkRows[index - 1] + 1,
  )
  expect(gapEnd).toBeGreaterThan(0)
  expect(inkRows[gapEnd] - inkRows[gapEnd - 1] - 1).toBeGreaterThanOrEqual(6)
})
