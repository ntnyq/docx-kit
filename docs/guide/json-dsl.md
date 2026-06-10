# JSON DSL (renderDocx)

The `renderDocx()` function accepts a JSON-serializable schema — ideal for AI/LLM-driven generation, API integrations, or storing document templates as data.

## Basic Usage

```ts
import { renderDocx } from 'docx-kit'

const blob = await renderDocx({
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
}).toBlob()
```

## Schema Structure

```ts
interface DocxSchema<TStyles> {
  content: BlockNode<TStyles>[]    // Ordered block nodes
  page?: DocxKitConfig['page']     // Page dimensions
  styles?: TStyles                 // Named stylesheet
}
```

## All Node Types

### Heading Nodes

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

### Paragraph Nodes

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

### Image Nodes

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

> **Note:** `data` must be in the format expected by the runtime (Uint8Array, Blob, etc. — not a base64 string with the prefix). For base64 data URLs inside Node.js, use `dataUrlToUint8Array()` from `docx-kit/node` to pre-process.

### Table Nodes

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

### Page Break Nodes

```ts
{ type: 'pageBreak' }
```

### Plugin Nodes

```ts
{
  type: 'plugin',
  name: 'qrcode' | 'echarts' | string,  // registered plugin name
  options: Record<string, unknown>,      // plugin-specific options
  style?: DocxStyleRule,
}
```

> **Important:** Plugin nodes require the plugin to be registered via `.use()` first. In `renderDocx()`, you register plugins on the returned builder before exporting.

```ts
// renderDocx returns a DocxBuilder — register plugins on it:
const doc = renderDocx({ content: [...] })
  .use(qrcodePlugin())
  .use(echartsPlugin())

// now .plugin(...) nodes will be resolved
await doc.save('output.docx')
```

## Complete JSON Example

```ts
import { renderDocx, qrcodePlugin, echartsPlugin } from 'docx-kit'

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

// Build + register plugins + export
await renderDocx(report)
  .use(echartsPlugin())
  .use(qrcodePlugin())
  .save('monthly-report.docx')
```

## AI/LLM Integration

The JSON DSL is designed for AI-driven document generation. Feed the schema structure to an LLM and it can generate complete `.docx` files:

```ts
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

const doc = renderDocx(filled)
await doc.save('ai-report.docx')
```
