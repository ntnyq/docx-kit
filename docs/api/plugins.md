# Plugin Types

Plugin system types and all built-in plugin option types.

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

Passed as the second argument to `render()`:

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

Type-safe plugin factory:

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

`context.style` contains the invocation's resolved stylesheet classes, theme
tokens, and inline overrides. All built-in plugins apply these overrides to
their paragraphs, text runs, and cells. Third-party plugins opt in through
`createPluginStyle(context.style)`, which provides `paragraph()`, `text()`, and
`cell()` adapters for native `docx` constructor options. Unspecified properties
retain the plugin's defaults, including nested spacing and borders.

```ts
doc.use(myPlugin).plugin('myPlugin', { text: 'Aligned label' }, {
  color: '#245A96',
  fontSize: 14,
  textAlign: 'right',
})
```

### `createPluginStyle()`

`createPluginStyle(style, options?)` accepts `PluginStyleOptions`: `defaults` and `overrides` are ordered arrays of optional `DocxStyleRule` layers; `theme` resolves tokens in those layers. Precedence is native plugin defaults → `defaults` → `style` → `overrides`. Nested fields merge, arrays replace, and `undefined` retains existing values. Compile layers separately so shorthand and aliases respect precedence.

```ts
const titleStyle = createPluginStyle(context.style, {
  defaults: [pluginConfig.styles?.title],
  overrides: [options.styles?.title],
  theme: context.config.theme,
})
const run = new TextRun(titleStyle.text({ text: options.title }))
```


Part styles use the same rules in section headers and footers, including `default`, `first`, and `even`. The AI tool schema, MCP schema resource, and runtime validator all check these nested plugin options.

A border color or width can be specified alone: existing border attributes are retained, and a new border receives the compiler's default `single` style and 1/8 pt width. A later `fontFamilyEastAsia` preserves an earlier Latin font. The `docx` typography escape hatch participates in the same cascade, including explicit `bold: false` and `italics: false`.

Margin shorthand can mix lengths and theme tokens, for example `margin: '$spacing.sm $spacing.md'`. Numeric spacing tokens use points.

---

## Built-in Plugin Option Types

All 25 built-in plugins and their option types.

Every built-in factory accepts an optional `<Plugin>Config` with default `styles`; each invocation accepts its own typed `styles`. See [Callout](/plugins/callout#style-customization) and [DataTable](/plugins/data-table#style-customization) for part styles, precedence and JSON DSL examples.

### `CrossReferenceOptions`

Factories, options, defaults and complete examples: [Cross Reference](../plugins/cross-reference.md).

`crossReferencePlugin` / `CrossReferenceConfig` / `CrossReferenceStyles`

### `BibliographyOptions`

Factories, options, defaults and complete examples: [Bibliography](../plugins/bibliography.md).

`bibliographyPlugin` / `BibliographyConfig` / `BibliographyStyles` / `BibliographyEntry` / `BibliographyBaseOptions` / `BibliographyListOptions` / `BibliographyCitationOptions`

### `GlossaryOptions`

Factories, options, defaults and complete examples: [Glossary](../plugins/glossary.md).

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

> Barcode rendering requires the optional `bwip-js` peer dependency.

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

> ECharts requires the `echarts` peer dependency. Works in browser environments; Node.js needs a server-side canvas.

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

> Image input and sizing follow the core image compiler. File paths require `docx-kit/node` or `config.resolveImage`; URLs require an explicit resolver.

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

> `change` and `trend` work independently. The final row is padded; empty input renders a placeholder.

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

> QR Code requires the `qrcode` peer dependency.

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

## Plugin Quick Reference

| Plugin                 | Node Name        | Output                       | Async? | Peer Deps                 |
| ---------------------- | ---------------- | ---------------------------- | ------ | ------------------------- |
| `badgePlugin`          | `badge`          | Paragraph                    | No     | —                         |
| `barcodePlugin`        | `barcode`        | Paragraph(s)                 | Yes    | `bwip-js`                 |
| `calloutPlugin`        | `callout`        | Table                        | No     | —                         |
| `changelogPlugin`      | `changelog`      | [Paragraph, Table]           | No     | —                         |
| `codeBlockPlugin`      | `codeBlock`      | Paragraph[]                  | Yes    | `highlight.js` (optional) |
| `coverPagePlugin`      | `coverPage`      | Paragraph[]                  | No     | —                         |
| `dataTablePlugin`      | `dataTable`      | Table                        | No     | —                         |
| `dividerPlugin`        | `divider`        | Paragraph                    | No     | —                         |
| `echartsPlugin`        | `echarts`        | Paragraph(s)                 | Yes    | `echarts`                 |
| `invoicePlugin`        | `invoice`        | Paragraph/Table[]            | No     | —                         |
| `letterheadPlugin`     | `letterhead`     | Paragraph[]                  | No     | —                         |
| `meetingMinutesPlugin` | `meetingMinutes` | [H1, P, Table]               | No     | —                         |
| `pageNumberPlugin`     | `pageNumber`     | Paragraph                    | No     | —                         |
| `propertyTablePlugin`  | `propertyTable`  | Table                        | No     | —                         |
| `qrcodePlugin`         | `qrcode`         | Paragraph(s)                 | Yes    | `qrcode`                  |
| `crossReferencePlugin` | `crossReference` | Paragraph | No | — |
| `bibliographyPlugin` | `bibliography` | Paragraph[] | No | — |
| `glossaryPlugin` | `glossary` | (Paragraph \| Table)[] | No | — |
| `checklistPlugin`     | `checklist`      | Paragraph[]                  | No     | —                         |
| `figurePlugin`        | `figure`         | Paragraph[]                  | Yes    | —                         |
| `metricCardPlugin`    | `metricCard`     | Table or Paragraph           | No     | —                         |
| `signatureBlockPlugin` | `signatureBlock` | Table                        | No     | —                         |
| `timelinePlugin`       | `timeline`       | Table                        | No     | —                         |
| `tocPlugin`            | `toc`            | [Paragraph, TableOfContents] | No     | —                         |
| `watermarkPlugin`      | `watermark`      | Paragraph                    | No     | —                         |

## Other built-in part styles

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `<Plugin>Config` and `<Plugin>Styles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Plugin | Config / Styles | TextStyleRule | ParagraphStyleRule | CellStyleRule (without height) |
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

Only the explicit title is styled here. Word generates TOC entries using its TOC document styles when updating the field. ECharts remains browser-only. Chart content and dimensions use `option`, `width`, and `height`; these styles affect the surrounding paragraphs and caption.

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
