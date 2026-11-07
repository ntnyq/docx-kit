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

### `createPluginStyle()`

`createPluginStyle(style, options?)` 接受 `PluginStyleOptions`：`defaults` 與 `overrides` 是依序套用的可選 `DocxStyleRule` 層陣列，`theme` 用於解析各層中的主題 token。優先順序為原生外掛預設值 → `defaults` → `style` → `overrides`。巢狀欄位合併，陣列整體取代，`undefined` 沿用既有值。各層分別編譯，確保簡寫與別名遵循優先順序。

```ts
const titleStyle = createPluginStyle(context.style, {
  defaults: [pluginConfig.styles?.title],
  overrides: [options.styles?.title],
  theme: context.config.theme,
})
const run = new TextRun(titleStyle.text({ text: options.title }))
```


部位樣式在頁首、頁尾的 `default`、`first`、`even` 中遵循相同規則；AI 工具 schema、MCP schema 資源及執行階段驗證皆檢查這些巢狀外掛選項。

框線可僅指定色彩或寬度：既有框線屬性會保留，新建框線補齊編譯器預設的 `single` 樣式及 1/8 pt 線寬。後續設定 `fontFamilyEastAsia` 會保留既有拉丁字型。`docx` 文字排版直通也遵循相同優先順序，包括明確的 `bold: false` 及 `italics: false`。

Margin 簡寫可混合長度與主題 token，例如 `margin: '$spacing.sm $spacing.md'`；數值型間距 token 以 pt 解讀。

---

## 內建外掛選項型別

全部 25 個內建外掛及其選項型別。

所有內建工廠都支援可選的 `<Plugin>Config` 預設 `styles`，單次呼叫也支援對應的型別安全部位 `styles`。部位樣式、優先順序及 JSON DSL 範例見 [Callout](/zh-TW/plugins/callout#樣式自訂) 和 [DataTable](/zh-TW/plugins/data-table#樣式自訂)。

### `CrossReferenceOptions`

工廠、選項、預設值與完整範例：[交叉參照](../plugins/cross-reference.md)。

`crossReferencePlugin` / `CrossReferenceConfig` / `CrossReferenceStyles`

### `BibliographyOptions`

工廠、選項、預設值與完整範例：[參考文獻](../plugins/bibliography.md)。

`bibliographyPlugin` / `BibliographyConfig` / `BibliographyStyles` / `BibliographyEntry` / `BibliographyBaseOptions` / `BibliographyListOptions` / `BibliographyCitationOptions`

### `GlossaryOptions`

工廠、選項、預設值與完整範例：[術語與縮寫表](../plugins/glossary.md)。

`glossaryPlugin` / `GlossaryConfig` / `GlossaryStyles` / `GlossaryEntry`

### `BadgeOptions`

```ts
interface BadgeOptions {
  styles?: BadgeStyles
  text: string
  backgroundColor?: string
  color?: 'danger' | 'info' | 'neutral' | 'success' | 'warning' | string
}
```

### `BarcodeOptions`

```ts
interface BarcodeOptions {
  styles?: BarcodeStyles
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
  styles?: CalloutStyles
}
interface CalloutConfig {
  styles?: CalloutStyles
}

interface CalloutStyles {
  cell?: Omit<CellStyleRule, 'height'>
  paragraph?: ParagraphStyleRule
  title?: TextStyleRule
  content?: TextStyleRule
}

```

### `ChangelogOptions`

```ts
interface ChangelogOptions {
  styles?: ChangelogStyles
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
  styles?: ChecklistStyles
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
  styles?: CodeBlockStyles
  code: string
  language?: string // requires highlight.js peer dep
  showLineNumbers?: boolean
}
```

### `CoverPageOptions`

```ts
interface CoverPageOptions {
  styles?: CoverPageStyles
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
  styles?: DataTableStyles
}

type ColFormat = 'currency' | 'date' | 'number' | 'percent'
type ColAlign = 'left' | 'center' | 'right'
interface DataTableConfig {
  styles?: DataTableStyles
}

interface DataTableStyles {
  headerCell?: Omit<CellStyleRule, 'height'>
  bodyCell?: Omit<CellStyleRule, 'height'>
  stripedCell?: Omit<CellStyleRule, 'height'>
  headerParagraph?: ParagraphStyleRule
  bodyParagraph?: ParagraphStyleRule
  emptyParagraph?: ParagraphStyleRule
  headerText?: TextStyleRule
  bodyText?: TextStyleRule
  emptyText?: TextStyleRule
}

```

### `DividerOptions`

```ts
interface DividerOptions {
  styles?: DividerStyles
  color?: string
  spacingAfter?: number
  spacingBefore?: number
  style?: 'dashed' | 'dotted' | 'double' | 'solid'
}
```

### `EChartsPluginOptions`

```ts
interface EChartsPluginOptions {
  styles?: EChartsStyles
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
  styles?: FigureStyles
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
  styles?: InvoiceStyles
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
  styles?: LetterheadStyles
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
  styles?: MeetingMinutesStyles
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
  styles?: MetricCardStyles
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
  styles?: PageNumberStyles
  alignment?: 'left' | 'center' | 'right'
  fontSize?: number // half-points, default: 20 (10pt)
  showTotal?: boolean // "Page X of Y" when true
}
```

### `PropertyTableOptions`

```ts
interface PropertyTableOptions {
  styles?: PropertyTableStyles
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
  styles?: QRCodeStyles
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
  styles?: SignatureBlockStyles
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
  styles?: TimelineStyles
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
  styles?: TocStyles
  maxLevel?: number // clamped to 1–9, default: 3
  title?: string // default: 'Contents'
}
```

### `WatermarkOptions`

```ts
interface WatermarkOptions {
  styles?: WatermarkStyles
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
| `crossReferencePlugin` | `crossReference` | Paragraph | No | — |
| `bibliographyPlugin` | `bibliography` | Paragraph[] | No | — |
| `glossaryPlugin` | `glossary` | (Paragraph \| Table)[] | No | — |
| `checklistPlugin` | `checklist` | 段落陣列 | 否 | — |
| `figurePlugin` | `figure` | 段落陣列 | 是 | — |
| `metricCardPlugin` | `metricCard` | 表格或段落 | 否 | — |
| `signatureBlockPlugin` | `signatureBlock` | 表格 | 否 | — |
| `timelinePlugin` | `timeline` | 表格 | 否 | — |
| `tocPlugin` | `toc` | [段落, TableOfContents] | 否 | — |
| `watermarkPlugin` | `watermark` | 段落 | 否 | — |

## 其他內建外掛的部位樣式

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `<Plugin>Config` 和 `<Plugin>Styles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 外掛 | 設定 / 樣式型別 | 文字部位 | 段落部位 | 儲存格部位（不含 height） |
| --- | --- | --- | --- | --- |
| [crossReferencePlugin](../plugins/cross-reference.md) | `CrossReferenceConfig` / `CrossReferenceStyles` | text | paragraph | — |
| [bibliographyPlugin](../plugins/bibliography.md) | `BibliographyConfig` / `BibliographyStyles` | citation / entry / title | citationParagraph / entryParagraph / titleParagraph | — |
| [glossaryPlugin](../plugins/glossary.md) | `GlossaryConfig` / `GlossaryStyles` | definition / fullForm / group / headerText / term / title | bodyParagraph / groupParagraph / headerParagraph / titleParagraph | bodyCell / headerCell |
| [badgePlugin](../plugins/badge.md) | `BadgeConfig` / `BadgeStyles` | `text` | `paragraph` | — |
| [dividerPlugin](../plugins/divider.md) | `DividerConfig` / `DividerStyles` | — | `paragraph` | — |
| [pageNumberPlugin](../plugins/page-number.md) | `PageNumberConfig` / `PageNumberStyles` | `text` | `paragraph` | — |
| [watermarkPlugin](../plugins/watermark.md) | `WatermarkConfig` / `WatermarkStyles` | `text` | `paragraph` | — |
| [tocPlugin](../plugins/toc.md) | `TocConfig` / `TocStyles` | `title` | `titleParagraph` | — |
| [coverPagePlugin](../plugins/cover-page.md) | `CoverPageConfig` / `CoverPageStyles` | `author` / `date` / `organization` / `subtitle` / `title` | `authorParagraph` / `dateParagraph` / `organizationParagraph` / `ruleParagraph` / `subtitleParagraph` / `titleParagraph` | — |
| [letterheadPlugin](../plugins/letterhead.md) | `LetterheadConfig` / `LetterheadStyles` | `address` / `companyName` / `contact` / `tagline` | `addressParagraph` / `companyParagraph` / `contactParagraph` / `ruleParagraph` / `taglineParagraph` | — |
| [barcodePlugin](../plugins/barcode.md) | `BarcodeConfig` / `BarcodeStyles` | `caption` | `captionParagraph` / `imageParagraph` | — |
| [qrcodePlugin](../plugins/qrcode.md) | `QRCodeConfig` / `QRCodeStyles` | `caption` | `captionParagraph` / `imageParagraph` | — |
| [echartsPlugin](../plugins/echarts.md) | `EChartsConfig` / `EChartsStyles` | `caption` | `captionParagraph` / `imageParagraph` | — |
| [figurePlugin](../plugins/figure.md) | `FigureConfig` / `FigureStyles` | `caption` / `source` | `captionParagraph` / `imageParagraph` / `sourceParagraph` | — |
| [propertyTablePlugin](../plugins/property-table.md) | `PropertyTableConfig` / `PropertyTableStyles` | `emptyText` / `keyText` / `valueText` | `emptyParagraph` / `keyParagraph` / `valueParagraph` | `keyCell` / `stripedCell` / `valueCell` |
| [metricCardPlugin](../plugins/metric-card.md) | `MetricCardConfig` / `MetricCardStyles` | `change` / `emptyText` / `label` / `unit` / `value` | `changeParagraph` / `emptyParagraph` / `labelParagraph` / `paddingParagraph` / `valueParagraph` | `cell` |
| [signatureBlockPlugin](../plugins/signature-block.md) | `SignatureBlockConfig` / `SignatureBlockStyles` | `date` / `emptyText` / `label` / `signature` | `dateParagraph` / `emptyParagraph` / `labelParagraph` / `paddingParagraph` / `signatureParagraph` | `cell` |
| [changelogPlugin](../plugins/changelog.md) | `ChangelogConfig` / `ChangelogStyles` | `bodyText` / `emptyText` / `headerText` / `title` | `bodyParagraph` / `emptyParagraph` / `headerParagraph` / `titleParagraph` | `bodyCell` / `headerCell` |
| [meetingMinutesPlugin](../plugins/meeting-minutes.md) | `MeetingMinutesConfig` / `MeetingMinutesStyles` | `bodyText` / `headerText` / `metaLabel` / `metaText` / `title` | `bodyParagraph` / `headerParagraph` / `metaParagraph` / `titleParagraph` | `bodyCell` / `headerCell` |
| [invoicePlugin](../plugins/invoice.md) | `InvoiceConfig` / `InvoiceStyles` | `bodyText` / `headerText` / `metaText` / `notes` / `notesLabel` / `partyLabel` / `partyText` / `subtotal` / `tax` / `title` / `total` | `bodyParagraph` / `headerParagraph` / `notesParagraph` / `partyParagraph` / `titleParagraph` / `totalsParagraph` | `bodyCell` / `headerCell` |
| [checklistPlugin](../plugins/checklist.md) | `ChecklistConfig` / `ChecklistStyles` | `emptyText` / `group` / `item` / `note` / `owner` / `title` | `emptyParagraph` / `groupParagraph` / `itemParagraph` / `noteParagraph` / `titleParagraph` | — |
| [timelinePlugin](../plugins/timeline.md) | `TimelineConfig` / `TimelineStyles` | `connector` / `date` / `description` / `emptyText` / `title` | `connectorParagraph` / `contentParagraph` / `dateParagraph` / `emptyParagraph` | `connectorCell` / `contentCell` / `dateCell` |
| [codeBlockPlugin](../plugins/code-block.md) | `CodeBlockConfig` / `CodeBlockStyles` | `code` / `lineNumber` | `paragraph` | — |

此處僅自訂明確產生的標題。更新目錄欄位時，Word 使用文件的 TOC 樣式產生目錄項目。 ECharts 仍僅支援瀏覽器。圖表內容與尺寸使用 `option`、`width`、`height`，此處樣式控制周圍段落與圖說。

```ts
import { createDocx, metricCardPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(metricCardPlugin({
    styles: { value: { color: '#2563EB', fontSize: 24 } },
  }))
  .plugin('metricCard', {
    items: [{ label: 'Revenue', value: 128000, unit: 'USD' }],
    styles: { value: { bold: false }, cell: { margin: '8pt' } },
  }, { fontFamily: 'Arial' })
```
