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

All 12 built-in plugins and their option types.

### `CalloutOptions`

```ts
interface CalloutOptions {
  type: 'info' | 'success' | 'warning' | 'danger'
  content: string
  title?: string
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
| `calloutPlugin` | `callout` | Paragraph | No | — |
| `codeBlockPlugin` | `codeBlock` | Paragraph[] | Yes | `highlight.js` (optional) |
| `coverPagePlugin` | `coverPage` | Paragraph[] | No | — |
| `dataTablePlugin` | `dataTable` | Table | No | — |
| `echartsPlugin` | `echarts` | Paragraph(s) | Yes | `echarts` |
| `meetingMinutesPlugin` | `meetingMinutes` | [H1, P, Table] | No | — |
| `pageNumberPlugin` | `pageNumber` | Paragraph | No | — |
| `propertyTablePlugin` | `propertyTable` | Table | No | — |
| `qrcodePlugin` | `qrcode` | Paragraph(s) | Yes | `qrcode` |
| `signatureBlockPlugin` | `signatureBlock` | Table | No | — |
| `timelinePlugin` | `timeline` | Table | No | — |
| `watermarkPlugin` | `watermark` | Paragraph | No | — |
