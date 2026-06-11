# docx-kit

[![CI](https://github.com/ntnyq/docx-kit/workflows/CI/badge.svg)](https://github.com/ntnyq/docx-kit/actions)
[![NPM VERSION](https://img.shields.io/npm/v/docx-kit.svg)](https://www.npmjs.com/package/docx-kit)
[![NPM DOWNLOADS](https://img.shields.io/npm/dy/docx-kit.svg)](https://www.npmjs.com/package/docx-kit)
[![LICENSE](https://img.shields.io/github/license/ntnyq/docx-kit.svg)](https://github.com/ntnyq/docx-kit/blob/main/LICENSE)

**CSS-like DOCX API Kit** — Type-safe, plugin-extensible Word document generation for Node.js and the browser.

Built on top of [docx](https://github.com/dolanmiu/docx) (docxjs), docx-kit adds:

📖 **Documentation:** [https://docx-kit.ntnyq.dev](https://docx-kit.ntnyq.dev)

- **CSS-like stylesheet** with class names and cascade resolution
- **Plugin architecture** with type-safe registration and 12 built-in plugins
- **Fluent builder API** (h1–h6, p, table, image, pageBreak, bulletList, numberedList, hyperlink, section, plugin)
- **Unit system** with CSS unit support (pt, px, mm, cm, in, %)
- **Platform-specific entry points** for Node.js and browser

---

## Install

```shell
npm install docx-kit
```

```shell
yarn add docx-kit
```

```shell
pnpm add docx-kit
```

---

## Quick Start

```ts
import { createDocx, defineStyles } from 'docx-kit'

const styles = defineStyles({
  accent: { color: '#2563eb', fontWeight: 'bold' },
  muted: { color: '#6b7280', fontSize: 10 },
  highlight: { highlight: 'yellow' },
})

const doc = createDocx({ styles })
  .h1('Hello, docx-kit!')
  .p('This is a paragraph with default styling.')
  .p('This text is accented.', { className: 'accent' })
  .p('Highlighted text.', { className: 'highlight' })
  .bulletList([
    'Bullet lists',
    'Numbered lists',
    'Hyperlinks',
  ])
  .numberedList([
    'Install docx-kit',
    'Define your styles',
    'Generate your document',
  ])
  .hyperlink('https://github.com/ntnyq/docx-kit', 'Visit docx-kit')
  .table({
    columns: [
      { key: 'name', title: 'Name', width: '60%' },
      { key: 'value', title: 'Value', width: '40%' },
    ],
    data: [
      { name: 'Revenue', value: '$1.2M' },
      { name: 'Growth', value: '+15%' },
    ],
  })

await doc.save('report.docx')   // Node.js
// or
const blob = await doc.toBlob() // Browser

```

---

## API

### `createDocx(config?)` → `DocxBuilder`

The primary entry point. Returns a fluent builder for chaining content.

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx({
  page: {
    size: 'A4',
    orientation: 'landscape',
    margin: '2cm 2.5cm',
  },
  styles: { /* ... */ },
  defaults: {
    text: { fontFamily: 'Arial', fontSize: 11 },
    paragraph: { lineHeight: 1.5 },
  },
  metadata: {
    title: 'Annual Report',
    creator: 'Alice',
    subject: 'Finance',
    keywords: ['report', '2024'],
  },
})
```

### Builder Methods

| Method | Description |
|--------|-------------|
| `.h1(text, opts?)` | Level-1 heading |
| `.h2(text, opts?)` | Level-2 heading |
| `.h3(text, opts?)` | Level-3 heading |
| `.h4(text, opts?)` | Level-4 heading |
| `.h5(text, opts?)` | Level-5 heading |
| `.h6(text, opts?)` | Level-6 heading |
| `.p(text, opts?)` | Paragraph |
| `.bulletList(items, opts?)` | Bullet list |
| `.numberedList(items, opts?)` | Numbered list |
| `.hyperlink(url, text, opts?)` | Hyperlink |
| `.table({ columns, data, ... })` | Table with header row |
| `.image({ data, width?, height?, ... })` | Image (inline or floating) |
| `.pageBreak()` | Forced page break |
| `.section(config?)` | Start a new document section |
| `.plugin(name, options, style?)` | Invoke a registered plugin |
| `.use(plugin)` | Register a plugin |
| `.add(node)` | Add a raw DSL node |

### Export Methods

```ts
await doc.toBlob()        // → Blob (browser & Node.js)
await doc.toUint8Array()  // → Uint8Array
await doc.toBuffer()      // → Uint8Array (alias, NOT Node.js Buffer)
await doc.toBase64()      // → base64 string
await doc.save('f.docx')  // → writes to disk (Node.js only)
await doc.toDocument()    // → raw docx Document (for further manipulation)
doc.toJSON()              // → plain object (debug / AI serialization)
```

### DSL Nodes (via `.add()`)

| Node Type | Fields |
|-----------|--------|
| `{ type: 'heading', level: 1-6, text, className?, id?, style? }` | Heading |
| `{ type: 'paragraph', text?, children?, className?, id?, style? }` | Paragraph |
| `{ type: 'bulletList', items, bullet?, level?, className?, style? }` | Bullet list |
| `{ type: 'numberedList', items, numberingFormat?, start?, level?, className?, style? }` | Numbered list |
| `{ type: 'hyperlink', url, children, className?, style? }` | Hyperlink |
| `{ type: 'image', data, width?, height?, alt?, floating? }` | Image |
| `{ type: 'pageBreak' }` | Page break |
| `{ type: 'sectionBreak', config? }` | Section boundary (internal) |
| `{ type: 'table', columns, data, bordered?, striped?, header? }` | Table |
| `{ type: 'plugin', name, options }` | Plugin invocation |

### Table API

```ts
doc.table({
  columns: [
    { key: 'name', title: 'Name', width: '60%', align: 'left' },
    { key: 'value', title: 'Value', width: '40%', align: 'right' },
  ],
  data: [
    { name: 'Revenue', value: '$1.2M' },
    { name: 'Growth', value: '+15%' },
  ],
  bordered: true,
  striped: true,
  header: true,
  headerCellStyle: { fontWeight: 'bold' },
  cellStyle: { fontSize: 10 },
})
```

Advanced: custom renderers per column:

```ts
.columns([
  {
    key: 'value',
    title: 'Value',
    render: (val, row, idx) =>
      Number.parseFloat(val) > 0
        ? [{ type: 'text', text: `+${val}`, style: { color: '#16a34a' } }]
        : val,
  },
])
```

### Image API

```ts
doc.image({
  data: arrayBuffer,           // string | Uint8Array | ArrayBuffer | Blob
  width: 200,                   // px (bare number = px)
  height: 150,
  alt: 'A descriptive text',
  floating: {                   // optional
    wrap: 'square',
    x: '10pt',
    y: '20pt',
  },
})
```

### Style System

Define styles with `defineStyles()` and reference them via `className`:

```ts
import { defineStyles } from 'docx-kit'

const styles = defineStyles({
  accent: {
    color: '#2563eb',
    fontWeight: 'bold',
    fontSize: 14,
  },
  warning: {
    color: '#dc2626',
    backgroundColor: '#fef2f2',
    borderLeft: { color: '#dc2626', style: 'single', width: 3 },
  },
})
```

**Style properties** (38 total):

| Category | Properties |
|----------|------------|
| **Text** | `fontFamily`, `fontSize`, `fontWeight`, `fontStyle`, `color`, `underline`, `strike`, `allCaps`, `backgroundColor` |
| **Spacing** | `letterSpacing`, `lineHeight` |
| **Margins** | `margin`, `marginTop`, `marginRight`, `marginBottom`, `marginLeft` |
| **Padding** | `padding` |
| **Alignment** | `textAlign`, `textIndent`, `verticalAlign` |
| **Borders** | `border`, `borderTop`, `borderRight`, `borderBottom`, `borderLeft` |
| **Layout** | `width`, `height` |
| **Escape** | `docx` — merge directly into docxjs options |

Style cascade: `defaults → stylesheet class(es) → inline style`. Multiple class names are supported (`className: 'accent warning'` or `['accent', 'warning']`).

### BorderRule

```ts
type BorderRule = {
  color?: string    // hex color
  style?: 'dashed' | 'dotted' | 'double' | 'none' | 'single'
  width?: number    // pt (bare number = pt)
}
```

### Unit System

All size values accept CSS-like units:

```ts
type UnitValue = number | `${number}%` | `${number}cm` | `${number}in` | `${number}mm` | `${number}pt` | `${number}px`
```

Bare numbers are context-dependent: `pt` for fonts/spacing/borders, `px` for images.

### Page Configuration

```ts
type PageConfig = {
  size?: 'A3' | 'A4' | 'Legal' | 'Letter' | { width: UnitValue; height: UnitValue }
  orientation?: 'portrait' | 'landscape'
  margin?: UnitValue | [topBottom, leftRight] | [top, right, bottom, left]  // CSS shorthand
}
```

---

## Multi-Section Documents

Use `.section(config?)` to split a document into independent sections, each with
its own page size, orientation, margins, headers, and footers.

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx()
  .h1('Section 1 — A4 Portrait')
  .p('This is the first section.')

  // Start a new section in A3 landscape
  .section({ page: { size: 'A3', orientation: 'landscape' } })
  .h1('Section 2 — A3 Landscape')
  .p('This section uses a different page size.')

  // Section with custom headers and footers
  .section({
    header: {
      default: { children: ['Chapter 3', 'Confidential'] },
      first: { children: ['Title Page'] },
    },
    footer: {
      default: { children: ['Page 3'] },
    },
  })
  .h1('Section 3 — With Headers')
  .p('This section has headers and footers.')
```

### SectionConfig

```ts
type SectionConfig = {
  page?: PageConfig        // Section-level page size/orientation/margin override
  header?: HeaderFooterConfig
  footer?: HeaderFooterConfig
}
```

### HeaderFooterConfig

```ts
type HeaderFooterConfig = {
  default?: HeaderFooterContent  // Appears on all pages
  first?: HeaderFooterContent    // First page only (overrides default)
  even?: HeaderFooterContent     // Even pages only (overrides default)
}

type HeaderFooterContent = {
  children: string[]             // Each string → one Paragraph line
}
```

---

## Plugins

docx-kit ships with **12 built-in plugins**. Each plugin is registered via `.use()` and invoked with `.plugin(name, options)`.

| Plugin | Node Name | Description | Docs |
|--------|-----------|-------------|------|
| [Callout](#) | `callout` | Colored info / warning / success / danger boxes | [→](https://docx-kit.ntnyq.dev/plugins/callout) |
| [Code Block](#) | `codeBlock` | Syntax-highlighted code blocks with line numbers | [→](https://docx-kit.ntnyq.dev/plugins/code-block) |
| [Cover Page](#) | `coverPage` | Professional title page | [→](https://docx-kit.ntnyq.dev/plugins/cover-page) |
| [Data Table](#) | `dataTable` | Auto-inferred table from object arrays | [→](https://docx-kit.ntnyq.dev/plugins/data-table) |
| [ECharts](#) | `echarts` | ECharts charts as embedded images | [→](https://docx-kit.ntnyq.dev/plugins/echarts) |
| [Meeting Minutes](#) | `meetingMinutes` | Structured meeting notes with agenda table | [→](https://docx-kit.ntnyq.dev/plugins/meeting-minutes) |
| [Page Number](#) | `pageNumber` | Page number field for headers/footers | [→](https://docx-kit.ntnyq.dev/plugins/page-number) |
| [Property Table](#) | `propertyTable` | Key-value pair styled table | [→](https://docx-kit.ntnyq.dev/plugins/property-table) |
| [QR Code](#) | `qrcode` | QR code images from text or URLs | [→](https://docx-kit.ntnyq.dev/plugins/qrcode) |
| [Signature Block](#) | `signatureBlock` | Signature lines for contracts | [→](https://docx-kit.ntnyq.dev/plugins/signature-block) |
| [Timeline](#) | `timeline` | Chronological timeline as a styled table | [→](https://docx-kit.ntnyq.dev/plugins/timeline) |
| [Watermark](#) | `watermark` | Text watermark for document branding | [→](https://docx-kit.ntnyq.dev/plugins/watermark) |

📖 **Full plugin documentation:** [https://docx-kit.ntnyq.dev/plugins/](https://docx-kit.ntnyq.dev/plugins/)

### Quick Example

```ts
import { createDocx, calloutPlugin, dataTablePlugin } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .use(dataTablePlugin())
  .h1('Report')
  .plugin('callout', { type: 'info', content: 'System status: OK' })
  .plugin('dataTable', {
    data: [
      { name: 'Alice', role: 'Engineer' },
      { name: 'Bob', role: 'Designer' },
    ],
  })
```

### Custom Plugins

```ts
import { definePlugin } from 'docx-kit'

const myPlugin = definePlugin({
  name: 'signature' as const,
  setup: () => { /* one-time init */ },
  render: (options, ctx) => {
    return new ctx.constructor.Paragraph({ ... })
  },
})
```

---

## AI-Friendly Entry Point

Use `renderDocx()` to generate documents from JSON — ideal for AI / LLM output:

```ts
import { renderDocx } from 'docx-kit'

const blob = await renderDocx({
  content: [
    { type: 'heading', level: 1, text: 'AI-Generated Report' },
    { type: 'paragraph', text: 'Generated automatically from structured data.' },
    { type: 'table', columns: [
      { key: 'name', title: 'Name' },
      { key: 'value', title: 'Value' },
    ], data: [{ name: 'Foo', value: 'Bar' }] },
  ],
  styles: { accent: { color: '#2563eb', fontWeight: 'bold' } },
})
```

---

## Platform Entry Points

| Entry | Import | Purpose |
|-------|--------|---------|
| Main | `'docx-kit'` | Universal — Builder, styles, plugins, all types |
| Node.js | `'docx-kit/node'` | `saveDocument()` for file-system write |
| Browser | `'docx-kit/browser'` | `normalizeImageData()`, `dataUrlToUint8Array()` |

---

## Exports

### Functions

| Export | Signature |
|--------|-----------|
| `createDocx` | `(config?: DocxKitConfig) => DocxBuilder` |
| `renderDocx` | `(schema: DocxSchema) => Promise<Blob>` |
| `defineStyles` | `<T>(styles: T) => T` |
| `definePlugin` | `<N,O>(plugin: DocxPlugin<N,O>) => DocxPlugin<N,O>` |
| `qrcodePlugin` | `() => DocxPlugin<'qrcode', QRCodePluginOptions>` |
| `echartsPlugin` | `() => DocxPlugin<'echarts', EChartsPluginOptions>` |
| `calloutPlugin` | `() => DocxPlugin<'callout', CalloutOptions>` |
| `codeBlockPlugin` | `() => DocxPlugin<'codeBlock', CodeBlockOptions>` |
| `coverPagePlugin` | `() => DocxPlugin<'coverPage', CoverPageOptions>` |
| `dataTablePlugin` | `() => DocxPlugin<'dataTable', DataTableOptions>` |
| `meetingMinutesPlugin` | `() => DocxPlugin<'meetingMinutes', MeetingMinutesOptions>` |
| `pageNumberPlugin` | `() => DocxPlugin<'pageNumber', PageNumberOptions>` |
| `propertyTablePlugin` | `() => DocxPlugin<'propertyTable', PropertyTableOptions>` |
| `signatureBlockPlugin` | `() => DocxPlugin<'signatureBlock', SignatureBlockOptions>` |
| `timelinePlugin` | `() => DocxPlugin<'timeline', TimelineOptions>` |
| `watermarkPlugin` | `() => DocxPlugin<'watermark', WatermarkOptions>` |
| `dataUrlToUint8Array` | `(dataUrl: string) => Uint8Array \| Promise<Uint8Array>` |

### Types

| Type | Category |
|------|----------|
| `DocxBuilder`, `DocxSchema` | Builder |
| `BlockNode`, `HeadingNode`, `ParagraphNode`, `BulletListNode`, `NumberedListNode`, `HyperlinkNode`, `ImageNode`, `TableNode`, `PageBreakNode`, `SectionBreakNode`, `PluginNode`, `TextNode`, `BulletItem`, `InlineNode` | DSL |
| `DocxStyleRule`, `StyleSheet`, `BorderRule`, `BorderStyle`, `FontWeight`, `TextAlign`, `VerticalAlign` | Style |
| `DocxKitConfig`, `PageConfig`, `PageSize`, `Orientation`, `DocxTheme`, `SectionConfig`, `HeaderFooterConfig`, `HeaderFooterContent` | Config |
| `DocxPlugin`, `PluginRegistry`, `PluginRenderContext` | Plugin |
| `QRCodePluginOptions`, `EChartsPluginOptions`, `CalloutOptions`, `CodeBlockOptions`, `CoverPageOptions`, `DataTableOptions`, `ColAlign`, `ColFormat`, `PropertyTableOptions`, `PropertyItem`, `MeetingMinutesOptions`, `AgendaItem`, `PageNumberOptions`, `SignatureBlockOptions`, `SignatureParty`, `TimelineOptions`, `TimelineEvent`, `WatermarkOptions` | Plugin options |
| `ClassName`, `BaseNode`, `TableColumn` | Utility types |
| `DocxKitError`, `ERROR_CODES`, `ErrorCode` | Errors |
| `UnitValue` | Units |

---

## Errors

All errors extend `DocxKitError` with `code` and `cause`:

| Code | Description |
|------|-------------|
| `EXPORT_FAILED` | Packaging/export failed |
| `IMAGE_INVALID_DATA` | Image data is null or empty |
| `PLUGIN_NOT_REGISTERED` | Plugin name not found in register |
| `PLUGIN_RENDER_FAILED` | Plugin render() threw |
| `STYLE_UNKNOWN_CLASS` | className not found in stylesheet |
| `TABLE_INVALID_COLUMNS` | Table columns is empty |
| `UNKNOWN_NODE_TYPE` | Unrecognized node type |

---

## Project Structure

```
src/
├── builder/        # DocxBuilder class + createDocx/renderDocx factories
├── compiler/       # Node → docxjs object compilation
│   ├── compileDocument.ts   # Top-level document assembly
│   ├── compileNode.ts       # Node type dispatcher
│   ├── compileStyle.ts      # Style → docxjs option mapping
│   └── units.ts             # Unit conversion (CSS → twips)
├── dsl/
│   └── nodes.ts             # All node type definitions
├── style/
│   └── normalizeStyle.ts    # Style cascade resolution
├── types/
│   ├── document.ts          # DocxKitConfig, PageConfig, DocxTheme
│   ├── plugin.ts            # DocxPlugin, PluginRenderContext
│   ├── style.ts             # DocxStyleRule, BorderRule, etc.
│   └── utility.ts           # UnitValue, LiteralUnion, HexColor
├── renderer/
│   └── pack.ts              # Export wrappers (packToBlob, etc.)
├── plugins/
│   ├── callout/             # Built-in Callout plugin
│   ├── code-block/          # Built-in Code Block plugin
│   ├── cover-page/          # Built-in Cover Page plugin
│   ├── data-table/          # Built-in Data Table plugin
│   ├── echarts/             # Built-in ECharts plugin
│   ├── meeting-minutes/     # Built-in Meeting Minutes plugin
│   ├── page-number/         # Built-in Page Number plugin
│   ├── property-table/      # Built-in Property Table plugin
│   ├── qrcode/              # Built-in QRCode plugin
│   ├── signature-block/     # Built-in Signature Block plugin
│   ├── timeline/            # Built-in Timeline plugin
│   └── watermark/           # Built-in Watermark plugin
├── node/
│   ├── fs.ts                # saveDocument (Node.js only)
│   └── dataUrl.ts           # Node.js data URL decoder
├── browser/
│   └── dom.ts               # Browser helpers
├── utils/
│   ├── dataUrl.ts           # Shared data URL utilities
│   └── image.ts             # Image factory helpers
├── errors.ts                # DocxKitError + error codes
├── index.ts                 # Main entry point
├── node.ts                  # Node.js entry point
└── browser.ts               # Browser entry point
```

---

## Feature Gap Analysis

See [docs/reports/docx-kit-gap-analysis.md](./docs/reports/docx-kit-gap-analysis.md) for a comprehensive comparison with `dolanmiu/docx` v9.7.1. Key gaps:

| Priority | Status | Missing Features |
|----------|--------|-----------------|
| **P0** | ✅ Done | Numbering/bullet lists, numbered lists, multiple sections, headers & footers |
| **P1** | Open | Hyperlinks, text highlighting, super/subscript, cell merging, page numbers, paragraph borders, cell shading |
| **P2** | Open | Bookmarks, TOC, table style presets, tab stops, keep-with-next, page borders, RTL, character spacing |

---

## License

[MIT](./LICENSE) License © 2025-PRESENT [ntnyq](https://github.com/ntnyq)
