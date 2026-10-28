# 条形码（Barcode）

条形码插件使用 `bwip-js` 生成一维条形码图片，
并作为 PNG 内容嵌入文档。

```ts
import {
  barcodePlugin,
  type BarcodeFormat,
  type BarcodeOptions,
} from 'docx-kit'
```

## 安装

使用 docx-kit 时同时安装可选渲染器：

```shell
pnpm add bwip-js
```

## 选项

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
`interleaved2of5`、`isbn`、`itf14`、`upca` 与 `upce`。`format` 类型也
支持 `bwip-js` 提供的其他编码器名称。

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

## 颜色与旋转

```ts
doc.plugin('barcode', {
  text: 'DOCX-KIT-2026',
  backgroundColor: '#FEF3C7',
  barColor: '#1F2937',
  rotate: 'R',
  textColor: '#DC2626',
})
```

在 Node.js 中使用 PNG 渲染器，在浏览器构建中使用 canvas
渲染器。显示高度根据生成的图片计算，
因此改变 `width` 时会保持条形码宽高比。
