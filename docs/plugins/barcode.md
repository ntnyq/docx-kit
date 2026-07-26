# Barcode

The barcode plugin generates linear barcode images with `bwip-js` and embeds
them into the document as PNG content.

```ts
import {
  barcodePlugin,
  type BarcodeFormat,
  type BarcodeOptions,
} from 'docx-kit'
```

## Install

Install the optional renderer alongside docx-kit:

```shell
pnpm add bwip-js
```

## Options

```ts
interface BarcodeOptions {
  text: string
  alignment?: 'center' | 'left' | 'right'
  backgroundColor?: string
  barColor?: string
  barHeight?: number
  caption?: string
  format?: BarcodeFormat
  includeText?: boolean
  rotate?: 'I' | 'L' | 'N' | 'R'
  scale?: number
  textColor?: string
  width?: number
}
```

Common formats include `code128`, `code39`, `code93`, `ean13`, `ean8`,
`interleaved2of5`, `isbn`, `itf14`, `upca`, and `upce`. The `format` type also
accepts other encoder names supported by `bwip-js`.

## Code 128

```ts
import { barcodePlugin, createDocx } from 'docx-kit/node'

const doc = createDocx()
  .use(barcodePlugin())
  .plugin('barcode', {
    text: 'DOCX-KIT-2026',
    caption: 'Inventory identifier',
    format: 'code128',
    width: 280,
  })

await doc.save('barcode.docx')
```

## EAN-13

```ts
doc.plugin('barcode', {
  text: '5901234123457',
  format: 'ean13',
  barHeight: 16,
  includeText: true,
})
```

## Colors and Rotation

```ts
doc.plugin('barcode', {
  text: 'DOCX-KIT-2026',
  backgroundColor: '#FEF3C7',
  barColor: '#1F2937',
  rotate: 'R',
  textColor: '#DC2626',
})
```

The plugin uses the Node.js PNG renderer when available and the browser canvas
renderer in browser builds. The displayed height is calculated from the
generated image so changing `width` preserves the barcode aspect ratio.
