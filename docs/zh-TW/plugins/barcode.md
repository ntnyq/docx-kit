# Barcode 一維條碼

Barcode 外掛使用 `bwip-js` 產生一維條碼 PNG，並將其嵌入 Word 文件。

## 安裝

```shell
pnpm add bwip-js
```

## 使用

```ts
import { barcodePlugin, createDocx } from 'docx-kit/node'

const doc = createDocx()
  .use(barcodePlugin())
  .plugin('barcode', {
    text: 'DOCX-KIT-2026',
    caption: '庫存編號',
    format: 'code128',
    width: 280,
  })

await doc.save('barcode.docx')
```

常用 `format` 包含 `code128`、`code39`、`ean13`、`ean8`、`isbn`、`itf14`、
`upca` 與 `upce`。也可以透過 `barHeight`、`includeText`、`barColor`、
`backgroundColor`、`rotate` 和 `alignment` 控制輸出。
