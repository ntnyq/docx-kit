# QR Code

Generates QR code images from text or URLs and embeds them as inline images in the document.

## Import

```ts
import { qrcodePlugin, type QRCodePluginOptions } from 'docx-kit'
```

## Dependencies

Requires the `qrcode` package as a peer dependency:

```bash
pnpm add qrcode
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `text` | `string` | _(required)_ | The text or URL to encode |
| `size` | `number` | `128` | QR code image size in pixels |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | `'M'` | Error correction level |
| `margin` | `number` | `1` | Quiet zone margin in modules |
| `caption` | `string` | — | Optional caption text below the QR code |

### Error Correction Levels

| Level | Recovery | Use Case |
|-------|----------|----------|
| `L` | ~7% | High-density data, clean environment |
| `M` | ~15% | General purpose (default) |
| `Q` | ~25% | Reliable scanning needed |
| `H` | ~30% | Logos overlaid, damaged environments |

## Examples

### Basic QR Code

```ts
import { createDocx, qrcodePlugin } from 'docx-kit'

const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Website QR')
  .plugin('qrcode', {
    text: 'https://example.com',
  })
  .save('basic-qr.docx')
```

### Business Card QR

```ts
const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Business Card')
  .plugin('qrcode', {
    text: 'https://ntnyq.com',
    size: 256,
    errorCorrectionLevel: 'H',
    margin: 2,
    caption: 'Scan to visit my website',
  })
  .save('business-card.docx')
```

### Payment QR Code

```ts
const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Invoice #2026-001')
  .p('Please scan the QR code below to complete payment:')
  .plugin('qrcode', {
    text: 'https://pay.example.com/invoice/2026-001',
    size: 200,
    errorCorrectionLevel: 'M',
    caption: 'Pay with WeChat / Alipay',
  })
  .save('payment-qr.docx')
```

### Wi-Fi QR Code

```ts
const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Office Wi-Fi')
  .p('Scan the QR code to connect to the office network:')
  .plugin('qrcode', {
    // Wi-Fi QR code format
    text: 'WIFI:T:WPA;S:Office-Network;P:password123;;',
    size: 180,
    caption: 'Office Wi-Fi',
  })
  .save('wifi-qr.docx')
```

### Small QR Code with No Caption

```ts
const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Document Reference')
  .plugin('qrcode', {
    text: 'REF-2026-06-11-001',
    size: 80,
    margin: 0,
  })
  .save('small-qr.docx')
```

### Multiple QR Codes

```ts
const doc = createDocx()
  .use(qrcodePlugin)
  .h1('Multi-Channel Contact')

  .h2('Website')
  .plugin('qrcode', {
    text: 'https://example.com',
    caption: 'Visit our website',
  })

  .h2('WeChat')
  .plugin('qrcode', {
    text: 'https://weixin.qq.com/q/xxxxx',
    caption: 'Follow us on WeChat',
  })

  .h2('Email')
  .plugin('qrcode', {
    text: 'mailto:contact@example.com',
    caption: 'Send us an email',
  })

  .save('multi-qr.docx')
```
