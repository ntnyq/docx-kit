# Plugin Types

Plugin system types and all built-in plugin option types.

## `DocxPlugin<TName, TOptions>`

```ts
interface DocxPlugin<
  TName extends string = string,
  TOptions = unknown,
> {
  name: TName
  setup?: (context: PluginRenderContext) => MaybePromise<void>
  render: (
    options: TOptions,
    context: PluginRenderContext,
  ) => MaybePromise<unknown>
}
```

## `PluginRenderContext`

Passed as the second argument to `render()`:

```ts
interface PluginRenderContext {
  config: DocxKitConfig
  compileNode: (node: BlockNode) => Promise<unknown>
  utils: {
    image: {
      fromBlob: (blob: Blob) => Promise<Uint8Array>
      fromDataUrl: (dataUrl: string) => MaybePromise<Uint8Array>
    }
  }
}
```

## `PluginRegistry`

```ts
type PluginRegistry = Record<string, unknown>
```

## `definePlugin()`

Type-safe plugin factory:

```ts
import { definePlugin } from 'docx-kit'

const myPlugin = definePlugin<'myPlugin', { text: string }>({
  name: 'myPlugin',
  render(options, ctx) {
    // Full access to docx library via ctx.compileNode()
    return new ctx.constructor.Paragraph({ text: options.text })
  },
})
```

---

## Built-in Plugin Option Types

All 19 built-in plugins and their option types.

### `BadgeOptions`

```ts
interface BadgeOptions {
  text: string
  backgroundColor?: string
  color?: 'danger' | 'info' | 'neutral' | 'success' | 'warning' | string
}
```

### `BarcodeOptions`

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

> Barcode rendering requires the optional `bwip-js` peer dependency.

### `CalloutOptions`

```ts
interface CalloutOptions {
  type: 'info' | 'success' | 'warning' | 'danger'
  content: string
  title?: string
}
```

### `ChangelogOptions`

```ts
interface ChangelogOptions {
  entries: ChangelogEntry[]
  title?: string
}

interface ChangelogEntry {
  changes: string
  date: string
  type: 'added' | 'changed' | 'fixed' | 'removed'
  version: string
}
```

### `CodeBlockOptions`

```ts
interface CodeBlockOptions {
  code: string
  language?: string           // requires highlight.js peer dep
  showLineNumbers?: boolean
}
```

### `CoverPageOptions`

```ts
interface CoverPageOptions {
  title: string
  subtitle?: string
  author?: string
  date?: string
  organization?: string
  backgroundColor?: string
  alignment?: 'center' | 'left' | 'right'
  showRule?: boolean
}
```

### `DataTableOptions`

```ts
interface DataTableOptions {
  data: Record<string, unknown>[]
  labels?: Record<string, string>   // human-readable column labels
  format?: Record<string, ColFormat> // per-column formatting
  align?: Record<string, ColAlign>   // per-column alignment
  bordered?: boolean
  striped?: boolean
}

type ColFormat = 'currency' | 'date' | 'number' | 'percent'
type ColAlign = 'left' | 'center' | 'right'
```

### `DividerOptions`

```ts
interface DividerOptions {
  color?: string
  spacingAfter?: number
  spacingBefore?: number
  style?: 'dashed' | 'dotted' | 'double' | 'solid'
}
```

### `EChartsPluginOptions`

```ts
interface EChartsPluginOptions {
  option: EChartsOption         // full ECharts option object
  width?: number                // default: 640
  height?: number               // default: 360
  renderer?: 'canvas' | 'svg'   // default: 'canvas'
  imageType?: 'png' | 'svg'     // default: 'png'
  caption?: string
}
```

> ECharts requires the `echarts` peer dependency. Works in browser environments; Node.js needs a server-side canvas.

### `InvoiceOptions`

```ts
interface InvoiceOptions {
  date: string
  from: InvoiceParty
  invoiceNumber: string
  items: InvoiceLineItem[]
  to: InvoiceParty
  currency?: string
  dueDate?: string
  notes?: string
  taxRate?: number
}

interface InvoiceLineItem {
  description: string
  quantity: number
  unitPrice: number
}

interface InvoiceParty {
  name: string
  address?: string
  email?: string
}
```

### `LetterheadOptions`

```ts
interface LetterheadOptions {
  companyName: string
  address?: string
  email?: string
  phone?: string
  tagline?: string
  website?: string
}
```

### `MeetingMinutesOptions`

```ts
interface MeetingMinutesOptions {
  title: string
  date: string
  attendees: string[]
  agenda: AgendaItem[]
}

interface AgendaItem {
  topic: string
  discussion: string
  decision?: string
  owner?: string
}
```

### `PageNumberOptions`

```ts
interface PageNumberOptions {
  alignment?: 'left' | 'center' | 'right'
  fontSize?: number             // half-points, default: 20 (10pt)
  showTotal?: boolean           // "Page X of Y" when true
}
```

### `PropertyTableOptions`

```ts
interface PropertyTableOptions {
  items: PropertyItem[]
  keyBold?: boolean             // default: true
  striped?: boolean             // default: true
}

interface PropertyItem {
  key: string
  value: string
}
```

### `QRCodePluginOptions`

```ts
interface QRCodePluginOptions {
  text: string
  size?: number                 // default: 128
  errorCorrectionLevel?: 'L' | 'M' | 'Q' | 'H'
  margin?: number               // default: 1
  caption?: string
}
```

> QR Code requires the `qrcode` peer dependency.

### `SignatureBlockOptions`

```ts
interface SignatureBlockOptions {
  parties: SignatureParty[]
  columns?: number              // default: 2
}

interface SignatureParty {
  label: string
  name?: string
  date?: string
}
```

### `TimelineOptions`

```ts
interface TimelineOptions {
  events: TimelineEvent[]
  layout?: 'alternating' | 'left' | 'right'
  accentColor?: string          // default: '4472C4'
}

interface TimelineEvent {
  date: string
  title: string
  description?: string
}
```

### `TocOptions`

```ts
interface TocOptions {
  maxLevel?: number             // clamped to 1–9, default: 3
  title?: string                // default: 'Contents'
}
```

### `WatermarkOptions`

```ts
interface WatermarkOptions {
  text: string
  color?: string               // hex without #, default: 'BFBFBF'
  fontSize?: number            // half-points, default: 48 (24pt)
  alignment?: 'left' | 'center' | 'right'
}
```

---

## Plugin Quick Reference

| Plugin | Node Name | Output | Async? | Peer Deps |
|---|---|---|---|---|
| `badgePlugin` | `badge` | Paragraph | No | — |
| `barcodePlugin` | `barcode` | Paragraph(s) | Yes | `bwip-js` |
| `calloutPlugin` | `callout` | Paragraph | No | — |
| `changelogPlugin` | `changelog` | [Paragraph, Table] | No | — |
| `codeBlockPlugin` | `codeBlock` | Paragraph[] | Yes | `highlight.js` (optional) |
| `coverPagePlugin` | `coverPage` | Paragraph[] | No | — |
| `dataTablePlugin` | `dataTable` | Table | No | — |
| `dividerPlugin` | `divider` | Paragraph | No | — |
| `echartsPlugin` | `echarts` | Paragraph(s) | Yes | `echarts` |
| `invoicePlugin` | `invoice` | Paragraph/Table[] | No | — |
| `letterheadPlugin` | `letterhead` | Paragraph[] | No | — |
| `meetingMinutesPlugin` | `meetingMinutes` | [H1, P, Table] | No | — |
| `pageNumberPlugin` | `pageNumber` | Paragraph | No | — |
| `propertyTablePlugin` | `propertyTable` | Table | No | — |
| `qrcodePlugin` | `qrcode` | Paragraph(s) | Yes | `qrcode` |
| `signatureBlockPlugin` | `signatureBlock` | Table | No | — |
| `timelinePlugin` | `timeline` | Table | No | — |
| `tocPlugin` | `toc` | [Paragraph, TableOfContents] | No | — |
| `watermarkPlugin` | `watermark` | Paragraph | No | — |
