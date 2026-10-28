# JSON DSL（renderDocx）

`renderDocx()` 接收可 JSON 序列化的 schema，適合 AI/LLM 驅動的產生、API 整合，以及將文件範本作為資料存儲。

## 基本用法

```ts
import { renderDocx } from 'docx-kit'

const doc = await renderDocx({
  page: { size: 'A4', margin: '20mm' },
  styles: {
    h1: { fontSize: 24, fontWeight: 'bold' },
    p:  { fontSize: 12 },
  },
  content: [
    { type: 'heading',  level: 1, text: 'My Document', className: 'h1' },
    { type: 'paragraph', text: 'This document was generated from JSON.', className: 'p' },
    { type: 'pageBreak' },
    { type: 'heading',  level: 2, text: 'Section 2', className: 'h1' },
    { type: 'paragraph', text: 'More content...', className: 'p' },
  ],
})
const blob = await doc.toBlob()
```

## Schema 結構

```ts
interface DocxSchema<TStyles> {
  content: BlockNode<TStyles>[]    // Ordered block nodes
  page?: DocxKitConfig['page']     // Page dimensions
  styles?: TStyles                 // Named stylesheet
}
```

## 全部節點型別

### 標題節點

```ts
{
  type: 'heading',
  level: 1 | 2 | 3 | 4 | 5 | 6,
  text: string,
  className?: string,
  id?: string,
  style?: DocxStyleRule,
}
```

```ts
{ type: 'heading', level: 1, text: 'Title', className: 'title' }
{ type: 'heading', level: 2, text: 'Section' }
{ type: 'heading', level: 3, text: 'Sub-section', id: 's1' }
```

### 段落節點

```ts
{
  type: 'paragraph',
  text: string,
  className?: string,
  id?: string,
  style?: DocxStyleRule,
  children?: InlineNode[],  // rich content
}
```

```ts
// Simple
{ type: 'paragraph', text: 'A simple paragraph.', className: 'body' }

// Rich with inline children
{
  type: 'paragraph',
  children: [
    { type: 'text', text: 'Regular text ' },
    { type: 'text', text: 'bold text', style: { fontWeight: 'bold' } },
    { type: 'text', text: ' and back to normal.' },
  ],
}
```

### 圖片節點

```ts
{
  type: 'image',
  data: string | Uint8Array | ArrayBuffer | Blob,
  width?: UnitValue,
  height?: UnitValue,
  alt?: string,
  imageType?: 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp',
  floating?: boolean | { wrap?: string, x?: UnitValue, y?: UnitValue },
  className?: string,
  style?: DocxStyleRule,
}
```

`data` 可直接接收位元組、Blob 或 data URL。使用 `docx-kit/node` 的 `renderDocx()` 時，也支援本地檔案路徑與 `file:` URL。核心或瀏覽器呼叫方必須提供位元組、data URL，或在執行階段選項中提供顯式 `resolveImage` 適配器，不會自動獲取路徑對應的內容。

### 表格節點

```ts
{
  type: 'table',
  columns: TableColumn[],
  data: Record<string, unknown>[],
  bordered?: boolean,
  striped?: boolean,
  header?: boolean,
  headerCellStyle?: DocxStyleRule,
  cellStyle?: DocxStyleRule,
  className?: string,
  style?: DocxStyleRule,
}
```

```ts
{
  type: 'table',
  columns: [
    { key: 'name', title: 'Name' },
    { key: 'value', title: 'Value', align: 'right' },
    {
      key: 'trend',
      title: 'Trend',
      align: 'center',
      render: (val) => [
        { type: 'text', text: val as string, style: { color: '#22c55e' } },
      ],
    },
  ],
  data: [
    { name: 'Revenue', value: '$1.2M', trend: '↑ 15%' },
    { name: 'Costs', value: '$800K', trend: '↓ 3%' },
  ],
  bordered: true,
  headerCellStyle: { fontWeight: 'bold', backgroundColor: '#f5f5f5' },
}
```

### 分頁節點

```ts
{ type: 'pageBreak' }
```

### 外掛節點

```ts
{
  type: 'plugin',
  name: 'qrcode' | 'echarts' | string,  // registered plugin name
  options: Record<string, unknown>,      // plugin-specific options
  style?: DocxStyleRule,
}
```

> **重要：** 外掛節點要求在 schema 的 `plugins` 陣列中聲明外掛，或在匯出前透過 `.use()` 註冊。

```ts
import { echartsPlugin, qrcodePlugin, renderDocx } from 'docx-kit'

const doc = await renderDocx({ content: [...] })
doc
  .use(qrcodePlugin())
  .use(echartsPlugin())

const blob = await doc.toBlob()
```

## 完整 JSON 範例

```ts
import { echartsPlugin, qrcodePlugin, renderDocx } from 'docx-kit'

const report = {
  page: { size: 'A4', margin: '20mm 25mm' },
  styles: {
    title: { fontSize: 28, fontWeight: 'bold', color: '#1a1a2e', textAlign: 'center' },
    h2: { fontSize: 18, fontWeight: 'bold', color: '#333' },
    body: { fontSize: 12, lineHeight: 1.5 },
    small: { fontSize: 9, color: '#999' },
  },
  content: [
    { type: 'heading', level: 1, text: 'Monthly Report', className: 'title' },
    { type: 'paragraph', text: 'June 2026', className: 'small', style: { textAlign: 'center' } },

    { type: 'heading', level: 2, text: 'Executive Summary', className: 'h2' },
    { type: 'paragraph', text: 'Revenue exceeded projections by 12%. Key product lines showed strong growth across all regions.', className: 'body' },

    { type: 'heading', level: 2, text: 'Key Metrics', className: 'h2' },
    {
      type: 'table',
      columns: [
        { key: 'metric', title: 'Metric' },
        { key: 'value', title: 'Value', align: 'right' },
        { key: 'change', title: 'Change', align: 'center' },
      ],
      data: [
        { metric: 'Total Revenue', value: '$1.45M', change: '+12%' },
        { metric: 'New Customers', value: '847', change: '+23%' },
        { metric: 'Churn Rate', value: '2.1%', change: '-0.5%' },
        { metric: 'NPS Score', value: '72', change: '+5' },
      ],
      bordered: true,
      striped: true,
      headerCellStyle: { fontWeight: 'bold', backgroundColor: '#f8f9fa' },
    },

    { type: 'pageBreak' },

    { type: 'heading', level: 2, text: 'Revenue Chart', className: 'h2' },
    {
      type: 'plugin',
      name: 'echarts',
      options: {
        option: {
          title: { text: 'Monthly Revenue' },
          xAxis: { data: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'] },
          yAxis: {},
          series: [{ type: 'bar', data: [200, 220, 250, 240, 270, 290] }],
        },
        caption: 'Figure 1: Monthly revenue trend',
      },
    },

    { type: 'heading', level: 2, text: 'Contact', className: 'h2' },
    {
      type: 'plugin',
      name: 'qrcode',
      options: {
        text: 'https://example.com/contact',
        size: 128,
        caption: 'Scan to visit our website',
      },
    },
  ],
}

// Build, register plugins, and export in the browser
const doc = await renderDocx(report)
doc
  .use(echartsPlugin())
  .use(qrcodePlugin())

const url = URL.createObjectURL(await doc.toBlob())
const anchor = document.createElement('a')
anchor.href = url
anchor.download = 'monthly-report.docx'
anchor.click()
URL.revokeObjectURL(url)
```

## AI/LLM 整合

JSON DSL 專為 AI 驅動的文件產生設計。向 LLM 提供 schema 結構，即可產生完整 `.docx` 檔案：

```ts
import { renderDocx } from 'docx-kit/node'

// Example: AI generates this JSON from a natural language prompt
const aiResponse = {
  page: { size: 'A4' },
  content: [
    { type: 'heading', level: 1, text: '{{TITLE}}' },
    { type: 'paragraph', text: '{{CONTENT}}' },
  ],
}

// Template interpolation
function fillTemplate(schema: any, vars: Record<string, string>) {
  const json = JSON.stringify(schema)
  const filled = json.replace(/\{\{(\w+)\}\}/g, (_, key) => vars[key] ?? '')
  return JSON.parse(filled)
}

const filled = fillTemplate(aiResponse, {
  TITLE: 'Q2 Earnings Report',
  CONTENT: 'Revenue grew 15% to $1.45M, exceeding analyst expectations.',
})

const doc = await renderDocx(filled)
await doc.save('ai-report.docx')
```
