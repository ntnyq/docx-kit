# 條碼（Barcode）

條碼外掛使用 `bwip-js` 產生一維條碼圖片，
並作為 PNG 內容嵌入文件。

```ts
import {
  barcodePlugin,
  type BarcodeFormat,
  type BarcodeOptions,
} from 'docx-kit'
```

## 安裝

使用 docx-kit 時同時安裝可選呈現器：

```shell
pnpm add bwip-js
```

## 選項

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

常用格式包括 `code128`、`code39`、`code93`、`ean13`、`ean8`、
`interleaved2of5`、`isbn`、`itf14`、`upca` 與 `upce`。`format` 型別也
支援 `bwip-js` 提供的其他編碼器名稱。

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

## 顏色與旋轉

```ts
doc.plugin('barcode', {
  text: 'DOCX-KIT-2026',
  backgroundColor: '#FEF3C7',
  barColor: '#1F2937',
  rotate: 'R',
  textColor: '#DC2626',
})
```

在 Node.js 中使用 PNG 呈現器，在瀏覽器建置中使用 canvas
呈現器。顯示高度根據產生的圖片計算，
因此改變 `width` 時會保持條碼寬高比。
