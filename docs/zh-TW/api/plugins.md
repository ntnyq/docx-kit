# 外掛型別

外掛系統型別與全部內建外掛的選項型別。

## `DocxPlugin<TName, TOptions>`

```ts
interface DocxPlugin<TName extends string = string, TOptions = unknown> {
  name: TName
  setup?: () => MaybePromise<void>
  render: (
    options: TOptions,
    context: PluginRenderContext,
  ) => MaybePromise<unknown>
}
```

## `PluginRenderContext`

作為第二個參數傳給 `render()`：

```ts
interface PluginRenderContext {
  config: DocxKitConfig
  style?: DocxStyleRule
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

型別安全的外掛工廠函式：

```ts
import { Paragraph } from 'docx'
import { createPluginStyle, definePlugin } from 'docx-kit'

const myPlugin = definePlugin<'myPlugin', { text: string }>({
  name: 'myPlugin',
  render(options, ctx) {
    const styles = createPluginStyle(ctx.style)
    return new Paragraph(styles.paragraph({ text: options.text }))
  },
})
```

`context.style` 包含本次呼叫解析後的樣式表類、主題
令牌和行內覆蓋。所有內建外掛都會將這些覆蓋應用於
段落、文字片段與儲存格。第三方外掛可透過
`createPluginStyle(context.style)` 接入，其 `paragraph()`、`text()` 和
`cell()` 適配器可產生原生 `docx` 構造選項。未指定的屬性
保留外掛預設值，包括嵌套的間距和邊框。

```ts
doc.use(myPlugin).plugin('myPlugin', { text: 'Aligned label' }, {
  color: '#245A96',
  fontSize: 14,
  textAlign: 'right',
})
```

---

## 內建外掛選項型別

全部 22 個內建外掛及其選項型別。

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

> 條碼呈現需要可選的 `bwip-js` 同儕相依套件。

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

### `ChecklistOptions`

```ts
interface ChecklistOptions {
  items: ChecklistItem[]
  title?: string
}

interface ChecklistItem {
  text: string
  checked?: boolean // default: false; native Word checkbox
  group?: string // heading on each group change; preserves input order
  note?: string
  owner?: string
}
```

### `CodeBlockOptions`

```ts
interface CodeBlockOptions {
  code: string
  language?: string // requires highlight.js peer dep
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
  showRule?: boolean // default: false; optional decorative rule
}
```

### `DataTableOptions`

```ts
interface DataTableOptions {
  data: Record<string, unknown>[]
  labels?: Record<string, string> // human-readable column labels
  format?: Record<string, ColFormat> // per-column formatting
  align?: Record<string, ColAlign> // per-column alignment
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
  option: EChartsOption // full ECharts option object
  width?: number // default: 640
  height?: number // default: 360
  renderer?: 'canvas' | 'svg' // default: 'canvas'
  imageType?: 'png' | 'svg' // default: 'png'
  caption?: string
}
```

> ECharts 需要 `echarts` 同儕相依套件，可用於瀏覽器環境；Node.js 需要伺服器端 canvas。

### `FigureOptions`

```ts
interface FigureOptions {
  data: string | ArrayBuffer | Blob | Uint8Array
  alignment?: 'center' | 'left' | 'right' // default: center
  alt?: string
  caption?: string // literal text; no automatic numbering
  height?: UnitValue
  imageType?: 'bmp' | 'gif' | 'jpeg' | 'jpg' | 'png'
  source?: string // literal credit; no prefix added
  width?: UnitValue
}
```

> 圖片輸入和尺寸遵循核心圖片編譯器的規則。本地路徑需要 `docx-kit/node` 或 `config.resolveImage`；URL 需要顯式解析器。

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

### `MetricCardOptions`

```ts
interface MetricCardOptions {
  items: MetricItem[]
  columns?: number // positive safe integer, default: 3
}

interface MetricItem {
  label: string
  value: number | string // displayed verbatim, including zero
  change?: string
  trend?: 'down' | 'neutral' | 'up'
  unit?: string
}
```

> `change` 與 `trend` 獨立工作。最後一行會補齊空儲存格；空輸入會呈現預留位置內容。

### `PageNumberOptions`

```ts
interface PageNumberOptions {
  alignment?: 'left' | 'center' | 'right'
  fontSize?: number // half-points, default: 20 (10pt)
  showTotal?: boolean // "Page X of Y" when true
}
```

### `PropertyTableOptions`

```ts
interface PropertyTableOptions {
  items: PropertyItem[]
  keyBold?: boolean // default: true
  striped?: boolean // default: true
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
  size?: number // default: 128
  errorCorrectionLevel?: 'L' | 'M' | 'Q' | 'H'
  margin?: number // default: 1
  caption?: string
}
```

> 二維碼需要 `qrcode` 同儕相依套件。

### `SignatureBlockOptions`

```ts
interface SignatureBlockOptions {
  parties: SignatureParty[]
  columns?: number // default: 2
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
  accentColor?: string // default: active theme accent, otherwise '2E75B6'
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
  maxLevel?: number // clamped to 1–9, default: 3
  title?: string // default: 'Contents'
}
```

### `WatermarkOptions`

```ts
interface WatermarkOptions {
  text: string
  color?: string // hex without #, default: 'BFBFBF'
  fontSize?: number // half-points, default: 48 (24pt)
  alignment?: 'left' | 'center' | 'right'
}
```

---

## 外掛速查

| 外掛 | 節點名稱 | 輸出 | 非同步？ | 同儕相依套件 |
| ---------------------- | ---------------- | ---------------------------- | ------ | ------------------------- |
| `badgePlugin` | `badge` | 段落 | 否 | — |
| `barcodePlugin` | `barcode` | 一個或多個段落 | 是 | `bwip-js` |
| `calloutPlugin` | `callout` | 表格 | 否 | — |
| `changelogPlugin` | `changelog` | [段落, 表格] | 否 | — |
| `codeBlockPlugin` | `codeBlock` | 段落陣列 | 是 | `highlight.js`（可選） |
| `coverPagePlugin` | `coverPage` | 段落陣列 | 否 | — |
| `dataTablePlugin` | `dataTable` | 表格 | 否 | — |
| `dividerPlugin` | `divider` | 段落 | 否 | — |
| `echartsPlugin` | `echarts` | 一個或多個段落 | 是 | `echarts` |
| `invoicePlugin` | `invoice` | 段落或表格陣列 | 否 | — |
| `letterheadPlugin` | `letterhead` | 段落陣列 | 否 | — |
| `meetingMinutesPlugin` | `meetingMinutes` | [一級標題, 段落, 表格] | 否 | — |
| `pageNumberPlugin` | `pageNumber` | 段落 | 否 | — |
| `propertyTablePlugin` | `propertyTable` | 表格 | 否 | — |
| `qrcodePlugin` | `qrcode` | 一個或多個段落 | 是 | `qrcode` |
| `checklistPlugin` | `checklist` | 段落陣列 | 否 | — |
| `figurePlugin` | `figure` | 段落陣列 | 是 | — |
| `metricCardPlugin` | `metricCard` | 表格或段落 | 否 | — |
| `signatureBlockPlugin` | `signatureBlock` | 表格 | 否 | — |
| `timelinePlugin` | `timeline` | 表格 | 否 | — |
| `tocPlugin` | `toc` | [段落, TableOfContents] | 否 | — |
| `watermarkPlugin` | `watermark` | 段落 | 否 | — |
