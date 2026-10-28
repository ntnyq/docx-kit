# Builder API

`DocxBuilder` 類提供鏈式 API，用於建置 Word 文件。

## 建立建構器

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

## 內容方法

### 標題（h1–h6）

全部六級標題都支援可選樣式覆蓋：

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

純文字或帶行內樣式的豐富內容：

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

### 行內豐富內容

透過 `span()` 與 `inlineImg()`，段落可以包含混合行內內容：

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

完整演示見[豐富內容範例](/zh-TW/examples/rich-content)。

### 圖片

嵌入 PNG、JPEG、GIF 或 BMP 圖片：

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

使用型別化欄建立資料表格：

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

### 分頁

插入強制分頁：

```ts
doc
  .h1('Chapter 1')
  .p('Content for chapter 1...')
  .pageBreak()
  .h1('Chapter 2')
  .p('Content for chapter 2...')
```

### 原始節點（add）

直接添加任意原始 DSL 節點：

```ts
doc.add({ type: 'heading', level: 2, text: 'Raw Heading' })
doc.add({ type: 'paragraph', text: 'Raw paragraph' })
doc.add({ type: 'pageBreak' })
```

### 外掛

呼叫已註冊外掛：

```ts
import { qrcodePlugin } from 'docx-kit/node'

doc
  .use(qrcodePlugin())     // Register QRCode
  .plugin('qrcode', { text: 'https://example.com', size: 200 })
```

詳見[外掛指南](/zh-TW/guide/plugins)。

## 匯出方法

| 方法 | 回傳值 | 平臺 | 使用場景 |
|---|---|---|---|
| `.save(filename)` | `Promise<void>` | Node.js | 將 .docx 包流式寫入磁盤 |
| `.toStream()` | `Promise<Readable>` | Node.js | 將 .docx 包流式輸出到其他目標 |
| `.toBlob()` | `Promise<Blob>` | 瀏覽器 | 透過 `URL.createObjectURL()` 下載 |
| `.toBuffer()` | `Promise<Uint8Array>` | 兩者 | 適用於任意場景的原始位元組 |
| `.toBase64()` | `Promise<string>` | 兩者 | HTTP 傳輸、資料庫存儲 |
| `.toDocument()` | `Promise<Document>` | 兩者 | 訪問原始 `docx` Document 實例 |
| `.toJSON()` | `object` | 兩者 | 序列化建構器狀態以便調試 |

### 瀏覽器下載範例

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

### Node.js 儲存範例

```ts
// In Node.js
import { createDocx } from 'docx-kit/node'

const doc = createDocx()
await doc.save('output.docx')

// Or pipe to another Node.js destination
const stream = await doc.toStream()
stream.pipe(uploadDestination)
```

## 外掛註冊（use）

先註冊外掛，再呼叫：

```ts
import { createDocx, qrcodePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(qrcodePlugin())     // Register QRCode
  .h1('Document with Plugins')
  .plugin('qrcode', { text: 'https://example.com' })
```

## 鏈式呼叫與可變性

所有內容方法都回傳 `this`，支援連續鏈式呼叫：

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

建構器是**可變的**，節點會累積到內部陣列中。若需要基於同一設定產生多份文件，請分別建立 `createDocx()` 實例。
