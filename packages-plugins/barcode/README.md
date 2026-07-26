# @docxkit/plugin-barcode

Linear barcode generation plugin for docx-kit.

## Usage

```ts
import { barcodePlugin } from '@docxkit/plugin-barcode'

const doc = createDocx()
  .use(barcodePlugin())
  .plugin('barcode', {
    format: 'code128',
    text: 'DOCX-KIT-2026',
  })
```
