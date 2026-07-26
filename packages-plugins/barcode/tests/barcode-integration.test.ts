import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { barcodePlugin } from '../src'

describe('barcodePlugin integration', () => {
  it('packs an actual bwip-js barcode as PNG media', async () => {
    const children = await barcodePlugin().render(
      {
        caption: 'Integration barcode',
        format: 'code128',
        text: 'DOCX-KIT-2026',
      },
      createPluginTestContext(),
    )
    if (
      !Array.isArray(children)
      || !children.every(child => child instanceof Paragraph)
    ) {
      throw new TypeError('Expected barcode paragraphs')
    }
    const document = new Document({
      sections: [{ children }],
    })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const mediaFile = Object.values(archive.files).find(
      file => !file.dir && file.name.startsWith('word/media/'),
    )

    expect(mediaFile).toBeDefined()
    if (!mediaFile) {
      throw new Error('Expected barcode media in generated package')
    }

    const media = await mediaFile.async('uint8array')
    expect([...media.slice(0, 8)]).toEqual([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    ])
  })
})
