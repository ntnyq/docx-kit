# Plugins

docx-kit currently ships with **18 built-in plugins**. Each plugin lives in its own workspace package under `packages-plugins/`, can be registered with `.use(...)`, and is invoked through `.plugin(name, options)`.

## Catalog

| Plugin | Package | Node Name | Description |
| --- | --- | --- | --- |
| [Badge](./badge) | `@docxkit/plugin-badge` | `badge` | Status chips and short labels |
| [Callout](./callout) | `@docxkit/plugin-callout` | `callout` | Info, warning, success, and danger boxes |
| [Changelog](./changelog) | `@docxkit/plugin-changelog` | `changelog` | Release-note tables with typed entries |
| [Code Block](./code-block) | `@docxkit/plugin-code-block` | `codeBlock` | Monospaced source blocks with optional highlighting |
| [Cover Page](./cover-page) | `@docxkit/plugin-cover-page` | `coverPage` | Professional title and report covers |
| [Data Table](./data-table) | `@docxkit/plugin-data-table` | `dataTable` | Object-array tables with formatting helpers |
| [Divider](./divider) | `@docxkit/plugin-divider` | `divider` | Horizontal rules and section separators |
| [ECharts](./echarts) | `@docxkit/plugin-echarts` | `echarts` | Embedded ECharts charts |
| [Invoice](./invoice) | `@docxkit/plugin-invoice` | `invoice` | Invoice layouts with totals and tax calculation |
| [Letterhead](./letterhead) | `@docxkit/plugin-letterhead` | `letterhead` | Branded company headers for formal letters |
| [Meeting Minutes](./meeting-minutes) | `@docxkit/plugin-meeting-minutes` | `meetingMinutes` | Structured meeting summaries |
| [Page Number](./page-number) | `@docxkit/plugin-page-number` | `pageNumber` | Current page or page X of Y fields |
| [Property Table](./property-table) | `@docxkit/plugin-property-table` | `propertyTable` | Key-value metadata tables |
| [QR Code](./qrcode) | `@docxkit/plugin-qrcode` | `qrcode` | QR code images from text or URLs |
| [Signature Block](./signature-block) | `@docxkit/plugin-signature-block` | `signatureBlock` | Signature grids for approvals and contracts |
| [Timeline](./timeline) | `@docxkit/plugin-timeline` | `timeline` | Chronological milestone rendering |
| [Table of Contents](./toc) | `@docxkit/plugin-toc` | `toc` | Word TOC fields based on headings |
| [Watermark](./watermark) | `@docxkit/plugin-watermark` | `watermark` | Simple textual watermarks |

## Shared Usage Pattern

```ts
import { createDocx, calloutPlugin } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'info',
    title: 'Heads up',
    content: 'This report is generated automatically.',
  })
```

## Compatibility Notes

| Plugin | Extra dependency or runtime note |
| --- | --- |
| `codeBlock` | Optional `highlight.js` peer dependency for syntax coloring |
| `echarts` | Requires `echarts`; browser DOM is the primary target |
| `qrcode` | Requires the `qrcode` package |
| `toc` | Word usually asks to update fields after opening the generated `.docx` |
| `pageNumber` | Most useful when inserted into headers or footers |

## How To Choose

- Use content plugins like `callout`, `codeBlock`, `badge`, and `divider` when you need presentational blocks inside a normal document flow.
- Use document-structure plugins like `coverPage`, `toc`, `pageNumber`, and `letterhead` when you want reusable report scaffolding.
- Use data plugins like `dataTable`, `propertyTable`, `timeline`, `invoice`, and `meetingMinutes` when your source data already has a clear schema.
- Use media plugins like `echarts` and `qrcode` when the document needs scannable or visual output.

Custom plugins are still supported through `definePlugin()`. See the [Plugins guide](/guide/plugins) and the [package catalog](/ecosystem/packages) for the package-level view.
