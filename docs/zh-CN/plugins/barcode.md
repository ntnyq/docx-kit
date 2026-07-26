# Barcode 条形码

Barcode 插件使用 `bwip-js` 生成一维条形码 PNG，并将其嵌入 Word 文档。

## 安装

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
    caption: '库存编号',
    format: 'code128',
    width: 280,
  })

await doc.save('barcode.docx')
```

常用 `format` 包括 `code128`、`code39`、`ean13`、`ean8`、`isbn`、`itf14`、
`upca` 与 `upce`。还可以通过 `barHeight`、`includeText`、`barColor`、
`backgroundColor`、`rotate` 和 `alignment` 控制输出。
