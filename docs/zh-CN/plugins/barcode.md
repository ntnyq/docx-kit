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

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `BarcodeConfig` 和 `BarcodeStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `caption` | `TextStyleRule` | 文字排版 |
| `captionParagraph`, `imageParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

这些样式控制 DOCX 段落和题注；条码像素、尺寸及图片内文字继续使用原有条码选项。

```ts
barcodePlugin({
  styles: { caption: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
