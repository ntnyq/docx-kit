# 插件类型

插件系统类型与全部内置插件的选项类型。

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

作为第二个参数传给 `render()`：

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

类型安全的插件工厂函数：

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

`context.style` 包含本次调用解析后的样式表类、主题
令牌和行内覆盖。所有内置插件都会将这些覆盖应用于
段落、文本片段与单元格。第三方插件可通过
`createPluginStyle(context.style)` 接入，其 `paragraph()`、`text()` 和
`cell()` 适配器可生成原生 `docx` 构造选项。未指定的属性
保留插件默认值，包括嵌套的间距和边框。

```ts
doc.use(myPlugin).plugin('myPlugin', { text: 'Aligned label' }, {
  color: '#245A96',
  fontSize: 14,
  textAlign: 'right',
})
```

### `createPluginStyle()`

`createPluginStyle(style, options?)` 接受 `PluginStyleOptions`：`defaults` 与 `overrides` 是按顺序应用的可选 `DocxStyleRule` 层数组，`theme` 用于解析各层中的主题 token。优先级为原生插件默认值 → `defaults` → `style` → `overrides`。嵌套字段合并，数组整体替换，`undefined` 沿用已有值。各层分别编译，保证简写与别名遵循优先级。

```ts
const titleStyle = createPluginStyle(context.style, {
  defaults: [pluginConfig.styles?.title],
  overrides: [options.styles?.title],
  theme: context.config.theme,
})
const run = new TextRun(titleStyle.text({ text: options.title }))
```


部位样式在页眉、页脚的 `default`、`first`、`even` 中遵循同样的规则；AI 工具 schema、MCP schema 资源和运行时校验均检查这些嵌套插件选项。

边框可仅指定颜色或宽度：已有边框属性会保留，新建边框补齐编译器默认的 `single` 样式和 1/8 pt 线宽。后续设置 `fontFamilyEastAsia` 会保留已有拉丁字体。`docx` 文字排版透传也遵循同样的优先级，包括显式的 `bold: false` 和 `italics: false`。

Margin 简写可混合长度和主题 token，例如 `margin: '$spacing.sm $spacing.md'`；数值型间距 token 按 pt 解释。

---

## 内置插件选项类型

全部 22 个内置插件及其选项类型。

所有内置工厂都支持可选的 `<Plugin>Config` 默认 `styles`，单次调用也支持对应的类型安全部位 `styles`。部位样式、优先级及 JSON DSL 示例见 [Callout](/zh-CN/plugins/callout#样式定制) 和 [DataTable](/zh-CN/plugins/data-table#样式定制)。

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

> 条形码渲染需要可选的 `bwip-js` 对等依赖。

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

> ECharts 需要 `echarts` 对等依赖，可用于浏览器环境；Node.js 需要服务端 canvas。

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

> 图片输入和尺寸遵循核心图片编译器的规则。本地路径需要 `docx-kit/node` 或 `config.resolveImage`；URL 需要显式解析器。

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

> `change` 与 `trend` 独立工作。最后一行会补齐空单元格；空输入会渲染占位内容。

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

> 二维码需要 `qrcode` 对等依赖。

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

## 插件速查

| 插件 | 节点名称 | 输出 | 异步？ | 对等依赖 |
| ---------------------- | ---------------- | ---------------------------- | ------ | ------------------------- |
| `badgePlugin` | `badge` | 段落 | 否 | — |
| `barcodePlugin` | `barcode` | 一个或多个段落 | 是 | `bwip-js` |
| `calloutPlugin` | `callout` | 表格 | 否 | — |
| `changelogPlugin` | `changelog` | [段落, 表格] | 否 | — |
| `codeBlockPlugin` | `codeBlock` | 段落数组 | 是 | `highlight.js`（可选） |
| `coverPagePlugin` | `coverPage` | 段落数组 | 否 | — |
| `dataTablePlugin` | `dataTable` | 表格 | 否 | — |
| `dividerPlugin` | `divider` | 段落 | 否 | — |
| `echartsPlugin` | `echarts` | 一个或多个段落 | 是 | `echarts` |
| `invoicePlugin` | `invoice` | 段落或表格数组 | 否 | — |
| `letterheadPlugin` | `letterhead` | 段落数组 | 否 | — |
| `meetingMinutesPlugin` | `meetingMinutes` | [一级标题, 段落, 表格] | 否 | — |
| `pageNumberPlugin` | `pageNumber` | 段落 | 否 | — |
| `propertyTablePlugin` | `propertyTable` | 表格 | 否 | — |
| `qrcodePlugin` | `qrcode` | 一个或多个段落 | 是 | `qrcode` |
| `checklistPlugin` | `checklist` | 段落数组 | 否 | — |
| `figurePlugin` | `figure` | 段落数组 | 是 | — |
| `metricCardPlugin` | `metricCard` | 表格或段落 | 否 | — |
| `signatureBlockPlugin` | `signatureBlock` | 表格 | 否 | — |
| `timelinePlugin` | `timeline` | 表格 | 否 | — |
| `tocPlugin` | `toc` | [段落, TableOfContents] | 否 | — |
| `watermarkPlugin` | `watermark` | 段落 | 否 | — |

## 其他内置插件的部位样式

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `<Plugin>Config` 和 `<Plugin>Styles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 插件 | 配置 / 样式类型 | 文本部位 | 段落部位 | 单元格部位（不含 height） |
| --- | --- | --- | --- | --- |
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

此处仅定制显式标题。更新目录域时，Word 使用文档的 TOC 样式生成目录条目。 ECharts 仍仅支持浏览器。图表内容与尺寸使用 `option`、`width`、`height`，这里的样式控制外围段落与题注。

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
