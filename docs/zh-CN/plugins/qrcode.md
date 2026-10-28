# 二维码（QR Code）

根据文字或 URL 生成二维码，并作为行内图片嵌入文档。

## 导入

```ts
import { qrcodePlugin, type QRCodePluginOptions } from 'docx-kit'
```

## 依赖

需要 `qrcode` 包作为对等依赖：

```bash
pnpm add qrcode
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `text` | `string` | _必填_ | 要编码的文字或 URL |
| `size` | `number` | `128` | 二维码图片尺寸，单位为像素 |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | `'M'` | 纠错级别 |
| `margin` | `number` | `1` | 静区边距，单位为模块 |
| `caption` | `string` | — | 二维码下方的可选题注 |

### 纠错级别

| 级别 | 恢复能力 | 使用场景 |
|-------|----------|----------|
| `L` | 约 7% | 高密度数据、干净环境 |
| `M` | 约 15% | 通用场景（默认） |
| `Q` | 约 25% | 需要可靠扫描 |
| `H` | 约 30% | 叠加标志或可能破损的环境 |

## 示例

### 基础二维码

```ts
import { createDocx, qrcodePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(qrcodePlugin())
  .h1('Website QR')
  .plugin('qrcode', {
    text: 'https://example.com',
  })
  .save('basic-qr.docx')
```

### 名片二维码

```ts
const doc = createDocx()
  .use(qrcodePlugin())
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

### 付款二维码

```ts
const doc = createDocx()
  .use(qrcodePlugin())
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

### Wi-Fi 二维码

```ts
const doc = createDocx()
  .use(qrcodePlugin())
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

### 无题注的小二维码

```ts
const doc = createDocx()
  .use(qrcodePlugin())
  .h1('Document Reference')
  .plugin('qrcode', {
    text: 'REF-2026-06-11-001',
    size: 80,
    margin: 0,
  })
  .save('small-qr.docx')
```

### 多个二维码

```ts
const doc = createDocx()
  .use(qrcodePlugin())
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
