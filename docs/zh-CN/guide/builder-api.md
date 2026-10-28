# Builder API

`DocxBuilder` 类提供链式 API，用于构建 Word 文档。

## 创建构建器

```ts
import { createDocx, defineStyles } from 'docx-kit/node'

const styles = defineStyles({
  title: { fontSize: 28, fontWeight: 'bold' },
  body: { fontSize: 12 },
})

const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm' },
})
```

## 内容方法

### 标题（h1–h6）

全部六级标题都支持可选样式覆盖：

```ts
doc
  .h1('Main Title', { className: 'title' })
  .h2('Section 1', { id: 'section-1', style: { color: '#333' } })
  .h3('Subsection 1.1')
  .h4('Details')
  .h5('Minor Heading')
  .h6('Smallest Heading')
```

### 段落

纯文本或带行内样式的富内容：

```ts
doc
  .p('A simple paragraph.')
  .p('Centered bold text', {
    className: 'body',
    style: { textAlign: 'center', fontWeight: 'bold' },
  })
  .p('With an explicit id', {
    id: 'summary',
    style: { lineHeight: 1.8 },
  })
```

### 行内富内容

通过 `span()` 与 `inlineImg()`，段落可以包含混合行内内容：

```ts
import { inlineImg, span } from 'docx-kit/node'

// Mixed text runs with per-span styles
doc.p([
  span('Hello '),
  span('world', { bold: true, color: '#e11d48' }),
  span(' — all in one paragraph.'),
])

// Inline image among text
doc.p([
  span('Click the icon: '),
  inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
  span(' to continue.'),
])
```

完整演示见[富内容示例](/zh-CN/examples/rich-content)。

### 图片

嵌入 PNG、JPEG、GIF 或 BMP 图片：

```ts
import { readFileSync } from 'node:fs'

const imageBytes = readFileSync('./logo.png')

doc.image({
  data: imageBytes,
  width: 200,
  height: 100,
  alt: 'Company Logo',
})

// With floating layout
doc.image({
  data: imageBytes,
  width: 150,
  height: 150,
  floating: {
    wrap: 'square',
    x: 100,
    y: 50,
  },
})
```

### 表格

使用类型化列创建数据表格：

```ts
interface SalesRow {
  product: string
  revenue: number
  growth: string
}

doc.table<SalesRow>({
  columns: [
    { key: 'product', title: 'Product' },
    { key: 'revenue', title: 'Revenue', align: 'right' },
    { key: 'growth', title: 'YoY Growth', align: 'center' },
  ],
  data: [
    { product: 'Widget A', revenue: 120000, growth: '+15%' },
    { product: 'Widget B', revenue: 85000, growth: '+8%' },
    { product: 'Widget C', revenue: 45000, growth: '-3%' },
  ],
  bordered: true,
  striped: true,
  headerCellStyle: {
    fontWeight: 'bold',
    backgroundColor: '#f0f0f0',
  },
})
```

### 分页

插入强制分页：

```ts
doc
  .h1('Chapter 1')
  .p('Content for chapter 1...')
  .pageBreak()
  .h1('Chapter 2')
  .p('Content for chapter 2...')
```

### 原始节点（add）

直接添加任意原始 DSL 节点：

```ts
doc.add({ type: 'heading', level: 2, text: 'Raw Heading' })
doc.add({ type: 'paragraph', text: 'Raw paragraph' })
doc.add({ type: 'pageBreak' })
```

### 插件

调用已注册插件：

```ts
import { qrcodePlugin } from 'docx-kit/node'

doc
  .use(qrcodePlugin())     // Register QRCode
  .plugin('qrcode', { text: 'https://example.com', size: 200 })
```

详见[插件指南](/zh-CN/guide/plugins)。

## 导出方法

| 方法 | 返回值 | 平台 | 使用场景 |
|---|---|---|---|
| `.save(filename)` | `Promise<void>` | Node.js | 将 .docx 包流式写入磁盘 |
| `.toStream()` | `Promise<Readable>` | Node.js | 将 .docx 包流式输出到其他目标 |
| `.toBlob()` | `Promise<Blob>` | 浏览器 | 通过 `URL.createObjectURL()` 下载 |
| `.toBuffer()` | `Promise<Uint8Array>` | 两者 | 适用于任意场景的原始字节 |
| `.toBase64()` | `Promise<string>` | 两者 | HTTP 传输、数据库存储 |
| `.toDocument()` | `Promise<Document>` | 两者 | 访问原始 `docx` Document 实例 |
| `.toJSON()` | `object` | 两者 | 序列化构建器状态以便调试 |

### 浏览器下载示例

```ts
// In the browser
const blob = await doc.toBlob()
const url = URL.createObjectURL(blob)
const a = document.createElement('a')
a.href = url
a.download = 'report.docx'
a.click()
URL.revokeObjectURL(url)
```

### Node.js 保存示例

```ts
// In Node.js
import { createDocx } from 'docx-kit/node'

const doc = createDocx()
await doc.save('output.docx')

// Or pipe to another Node.js destination
const stream = await doc.toStream()
stream.pipe(uploadDestination)
```

## 插件注册（use）

先注册插件，再调用：

```ts
import { createDocx, qrcodePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(qrcodePlugin())     // Register QRCode
  .h1('Document with Plugins')
  .plugin('qrcode', { text: 'https://example.com' })
```

## 链式调用与可变性

所有内容方法都返回 `this`，支持连续链式调用：

```ts
const doc = createDocx()

doc
  .h1('Title')
  .p('Paragraph 1')
  .p('Paragraph 2')
  .pageBreak()
  .h2('Section')
  .p('More content...')
  .save('output.docx')
```

构建器是**可变的**，节点会累积到内部数组中。若需要基于同一配置生成多份文档，请分别创建 `createDocx()` 实例。
