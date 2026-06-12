# @docxkit/plugin-watermark

Watermark plugin for docx-kit.

## Usage

```ts
import { watermarkPlugin } from '@docxkit/plugin-watermark'

builder.use(watermarkPlugin)
builder.watermark({ text: 'CONFIDENTIAL', opacity: 0.1 })
```
