# Builder API

The `DocxBuilder` class provides a fluent, chainable API for constructing Word documents.

## Creating a Builder

```ts
import { createDocx, defineStyles } from 'docx-kit'

const styles = defineStyles({
  title: { fontSize: 28, fontWeight: 'bold' },
  body: { fontSize: 12 },
})

const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm' },
})
```

## Content Methods

### Headings (h1–h6)

All six heading levels with optional style overrides:

```ts
doc
  .h1('Main Title', { className: 'title' })
  .h2('Section 1', { id: 'section-1', style: { color: '#333' } })
  .h3('Subsection 1.1')
  .h4('Details')
  .h5('Minor Heading')
  .h6('Smallest Heading')
```

### Paragraphs

Plain text or rich content with inline styling:

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

### Images

Embed PNG, JPEG, GIF, or BMP images:

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

### Tables

Create data tables with typed columns:

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

### Page Breaks

Insert forced page breaks:

```ts
doc
  .h1('Chapter 1')
  .p('Content for chapter 1...')
  .pageBreak()
  .h1('Chapter 2')
  .p('Content for chapter 2...')
```

### Raw Nodes (add)

Add any raw DSL node directly:

```ts
doc.add({ type: 'heading', level: 2, text: 'Raw Heading' })
doc.add({ type: 'paragraph', text: 'Raw paragraph' })
doc.add({ type: 'pageBreak' })
```

### Plugins

Invoke registered plugins:

```ts
doc
  .use(qrcodePlugin())
  .plugin('qrcode', { text: 'https://example.com', size: 200 })
```

See the [Plugins guide](/guide/plugins) for more details.

## Export Methods

| Method | Returns | Platform | Use Case |
|---|---|---|---|
| `.save(filename)` | `Promise<void>` | Node.js | Write .docx to disk |
| `.toBlob()` | `Promise<Blob>` | Browser | Download via `URL.createObjectURL()` |
| `.toBuffer()` | `Promise<Uint8Array>` | Both | Raw bytes for any use case |
| `.toBase64()` | `Promise<string>` | Both | HTTP transfer, DB storage |
| `.toDocument()` | `Promise<Document>` | Both | Access raw `docx` Document instance |
| `.toJSON()` | `object` | Both | Serialize builder state for debugging |

### Browser Download Example

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

### Node.js Save Example

```ts
// In Node.js
await doc.save('output.docx')

// Or manually
const bytes = await doc.toBuffer()
import { writeFileSync } from 'node:fs'
writeFileSync('output.docx', bytes)
```

## Plugin Registration (use)

Register plugins before invoking them:

```ts
import { createDocx, qrcodePlugin, echartsPlugin } from 'docx-kit'

const doc = createDocx()
  .use(qrcodePlugin())     // Register QRCode
  .use(echartsPlugin())    // Register ECharts
  .h1('Document with Plugins')
  .plugin('qrcode', { text: 'https://example.com' })
  .plugin('echarts', {
    option: {
      xAxis: { data: ['A', 'B', 'C'] },
      yAxis: {},
      series: [{ type: 'bar', data: [1, 2, 3] }],
    },
  })
```

## Chaining & Immutability

All content methods return `this`, enabling indefinite chaining:

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

The builder is **mutable** — nodes accumulate in the internal array. If you need multiple documents from the same config, create separate `createDocx()` instances.
