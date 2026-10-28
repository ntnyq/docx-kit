# 水印（Watermark）

通过带样式的段落渲染大号半透明文字水印，适合放在节的页眉或页脚。

## 导入

```ts
import { watermarkPlugin, type WatermarkOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `text` | `string` | _必填_ | 水印文字 |
| `alignment` | `AlignmentType` | `CENTER` | 水平对齐 |
| `color` | `string` | `'BFBFBF'` | RRGGBB 十六进制文字颜色 |
| `fontSize` | `number` | `48` | 字号，单位为半磅（48 = 24pt） |

## 示例

### 基础水印

```ts
import { createDocx, watermarkPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Confidential Document')
  .p('This document contains proprietary information.')
  .plugin('watermark', { text: 'CONFIDENTIAL' })
  .save('confidential.docx')
```

### 红色 DRAFT 水印

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Work in Progress')
  .p('This document is still under review.')
  .plugin('watermark', {
    text: 'DRAFT',
    color: 'FF0000',
    fontSize: 72, // 36pt
  })
  .save('draft.docx')
```

### 左对齐水印

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Internal Memo')
  .plugin('watermark', {
    text: 'INTERNAL USE ONLY',
    alignment: AlignmentType.LEFT,
    color: '999999',
  })
  .save('internal.docx')
```

### 放入节的页眉或页脚

要为整份文档添加水印，请放入页眉或页脚：

```ts
const doc = createDocx()
  .use(watermarkPlugin())

// Render watermark in a footer that applies to all pages
const wm = doc.plugin('watermark', {
  text: 'SAMPLE',
  color: 'D0D0D0',
  fontSize: 60,
})

// Note: The exact header/footer integration depends on
// how sections are configured in your document.
```

### 多个水印

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Review Document')
  .plugin('watermark', { text: 'DRAFT', color: 'FF6600' })
  .p('Please review the following content carefully.')
  .pageBreak()
  .plugin('watermark', { text: 'CONFIDENTIAL', color: 'CC0000' })
  .p('This section contains confidential financial data.')
  .save('multi-watermark.docx')
```
