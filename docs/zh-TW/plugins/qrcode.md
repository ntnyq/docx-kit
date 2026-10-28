# 二維碼（QR Code）

根據文字或 URL 產生二維碼，並作為行內圖片嵌入文件。

## 匯入

```ts
import { qrcodePlugin, type QRCodePluginOptions } from 'docx-kit'
```

## 相依

需要 `qrcode` 包作為同儕相依套件：

```bash
pnpm add qrcode
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `text` | `string` | _必填_ | 要編碼的文字或 URL |
| `size` | `number` | `128` | 二維碼圖片尺寸，單位為像素 |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | `'M'` | 糾錯級別 |
| `margin` | `number` | `1` | 靜區邊距，單位為模組 |
| `caption` | `string` | — | 二維碼下方的可選題注 |

### 糾錯級別

| 級別 | 恢復能力 | 使用場景 |
|-------|----------|----------|
| `L` | 約 7% | 高密度資料、乾淨環境 |
| `M` | 約 15% | 通用場景（預設） |
| `Q` | 約 25% | 需要可靠掃描 |
| `H` | 約 30% | 疊加標誌或可能破損的環境 |

## 範例

### 基本二維碼

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

### 名片二維碼

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

### 付款二維碼

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

### Wi-Fi 二維碼

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

### 無題注的小二維碼

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

### 多個二維碼

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
