# @docxkit/plugin-qrcode

QR code generation plugin for docx-kit.

## Usage

```ts
import { qrcodePlugin } from '@docxkit/plugin-qrcode'

builder.use(qrcodePlugin)
builder.qrcode({ text: 'https://example.com', size: 128 })
```
