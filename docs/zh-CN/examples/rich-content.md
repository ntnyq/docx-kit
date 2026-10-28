# 示例：富内容

展示段落中的行内富内容：混合文本片段、行内图片，以及 `span()`、`inlineImg()` DSL 辅助函数。

## 完整代码

```ts
import { createDocx, defineStyles, span, inlineImg } from 'docx-kit/node'

// 1. Styles
const styles = defineStyles({
  body: {
    fontSize: 12,
    lineHeight: 1.6,
  },
  caption: {
    fontSize: 10,
    color: '#888',
    textAlign: 'center',
    fontStyle: 'italic',
  },
})

// 2. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
})

// A small inline icon as a base64 data URL (1x1 px transparent gif placeholder)
const iconDataUrl =
  'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'

doc
  .h1('Rich Inline Content')
  .p('This example demonstrates mixed text runs and inline images within a paragraph.')

  // ─── Mixed text runs via children array ─────────────────────────────
  .p([
    span('This sentence has '),
    span('bold red text', { bold: true, color: '#e11d48' }),
    span(', '),
    span('italic blue text', { italic: true, color: '#2563eb' }),
    span(', and '),
    span('highlighted text', { highlight: 'yellow' }),
    span(' — all in the same paragraph.'),
  ])

  // ─── Inline image inside paragraph text ────────────────────────────
  .p([
    span('Here is an inline icon: '),
    inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
    span(' and some text after it.'),
  ])

  // ─── Styled link-like text ─────────────────────────────────────────
  .p([
    span('Visit our site: '),
    span('https://example.com', {
      color: '#2563eb',
      underline: true,
      bold: true,
    }),
    span(' for more info.'),
  ])

  // ─── Nested emphasis ──────────────────────────────────────────────
  .p([
    span('Note: '),
    span('This is important!', { bold: true, color: '#dc2626' }),
    span(' Please review before proceeding.'),
  ])

  .save('rich-content-demo.docx')
```

## 本例展示的能力

| 功能 | 使用位置 |
|---|---|
| `span(text, style?)` 辅助函数 | 创建带独立样式的行内文本片段 |
| `inlineImg(opts)` 辅助函数 | 将行内图片嵌入段落内容流 |
| 数组形式的 `ParagraphNode.children` | 在单个段落中混合内容 |
| 片段级 `bold`、`italic`、`color` | 段落内的行内样式覆盖 |
| 片段级 `highlight` | 段落中的文字高亮 |
| 片段级 `underline` | 段落中的下划线文字 |

## `span()` 辅助函数

```ts
import { span } from 'docx-kit'

// Plain text span
span('Hello ')

// Styled text span
span('world', { bold: true, color: '#f00' })

// Used inside p() children array
doc.p([
  span('Hello '),
  span('world', { bold: true }),
  span('!'),
])
```

## `inlineImg()` 辅助函数

```ts
import { inlineImg } from 'docx-kit'

// Inline image from URL, base64, Buffer, or Uint8Array
inlineImg({
  data: 'https://example.com/icon.png',
  width: 16,
  height: 16,
})

// Used inside p() children array
doc.p([
  span('See icon: '),
  inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
  span(' done!'),
])
```

## 段落子节点 API

```ts
// String (classic API — all text in one run)
doc.p('Hello world', { className: 'body' })

// TextNode array (fine-grained control)
doc.p([
  { text: 'Hello ', type: 'text' },
  { text: 'world', type: 'text', style: { bold: true } },
])

// Mixed: text + inline image
doc.p([
  { text: 'Icon: ', type: 'text' },
  { type: 'image', data: imgData, width: 16, height: 16 },
  { text: ' done!', type: 'text' },
])

// With span() / inlineImg() helpers (recommended — cleaner syntax)
doc.p([
  span('Icon: '),
  inlineImg({ data: imgData, width: 16, height: 16 }),
  span(' done!'),
])
```
