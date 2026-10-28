<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./docs/public/brand/logo-dark.svg">
    <img src="./docs/public/brand/logo-light.svg" alt="docx-kit" width="360" height="113">
  </picture>
</p>

<p align="center">
  <a href="https://github.com/ntnyq/docx-kit/actions"><img src="https://github.com/ntnyq/docx-kit/workflows/CI/badge.svg" alt="CI"></a>
  <a href="https://www.npmjs.com/package/docx-kit"><img src="https://img.shields.io/npm/v/docx-kit.svg" alt="NPM VERSION"></a>
  <a href="https://www.npmjs.com/package/docx-kit"><img src="https://img.shields.io/npm/dy/docx-kit.svg" alt="NPM DOWNLOADS"></a>
  <a href="https://github.com/ntnyq/docx-kit/blob/main/LICENSE"><img src="https://img.shields.io/github/license/ntnyq/docx-kit.svg" alt="LICENSE"></a>
</p>

<p align="center">
  English | <a href="./README.zh-CN.md">简体中文</a>
</p>

**CSS-like DOCX API Kit** — Type-safe, plugin-extensible Word document generation for Node.js & browser.

---

## 📖 Documentation

👉 **Full guides, API reference, plugins, and examples: https://docx-kit.ntnyq.dev**

---

## Features

- **CSS-like styling** — `fontSize`, `color`, `margin`, `border`... just like CSS
- **Fluent builder API** — `.h1()` `.p()` `.table()` `.image()` `.use().plugin()`
- **22 built-in plugins** — Barcodes, QR codes, ECharts, code blocks, invoices, TOCs, and more
- **Style presets** — `classic` (gov-doc), `modern` (business), `academic` (thesis)
- **Themes** — `minimal`, `ocean`, `warm` with token system (`$colors.primary`)
- **JSON DSL** — `renderDocx()` for AI/LLM-driven document generation
- **Markdown** — Convert Markdown to the existing DSL with `docx-kit/markdown`, reusing styles and plugins
- **AI templates & MCP server** — `report`, `invoice`, `resume`, `letter` + MCP tools
- **Cross-platform** — Node.js (`save()`) & browser (`toBlob()`)
- **Browser preview** — Render `.docx` files in the browser with `@docxkit/renderer`

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
import { createDocx, defineStyles } from 'docx-kit/node'

const doc = createDocx({
  styles: defineStyles({
    title: { fontSize: 24, fontWeight: 'bold', textAlign: 'center' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),
})

doc
  .h1('Hello, docx-kit!', { className: 'title' })
  .p('This is a paragraph.', { className: 'body' })
  .table({
    columns: [
      { key: 'name', title: 'Name' },
      { key: 'value', title: 'Value' },
    ],
    data: [{ name: 'Revenue', value: '$1.2M' }],
  })

await doc.save('output.docx') // Node.js entry
// Browser: import from 'docx-kit' and use await doc.toBlob()
```

## Plugins

22 built-in plugins:

- [Badge](https://docx-kit.ntnyq.dev/plugins/badge)
- [Barcode](https://docx-kit.ntnyq.dev/plugins/barcode)
- [Callout](https://docx-kit.ntnyq.dev/plugins/callout)
- [Changelog](https://docx-kit.ntnyq.dev/plugins/changelog)
- [Checklist](https://docx-kit.ntnyq.dev/plugins/checklist)
- [Code Block](https://docx-kit.ntnyq.dev/plugins/code-block)
- [Cover Page](https://docx-kit.ntnyq.dev/plugins/cover-page)
- [Data Table](https://docx-kit.ntnyq.dev/plugins/data-table)
- [Divider](https://docx-kit.ntnyq.dev/plugins/divider)
- [ECharts](https://docx-kit.ntnyq.dev/plugins/echarts)
- [Figure](https://docx-kit.ntnyq.dev/plugins/figure)
- [Invoice](https://docx-kit.ntnyq.dev/plugins/invoice)
- [Letterhead](https://docx-kit.ntnyq.dev/plugins/letterhead)
- [Meeting Minutes](https://docx-kit.ntnyq.dev/plugins/meeting-minutes)
- [Metric Card](https://docx-kit.ntnyq.dev/plugins/metric-card)
- [Page Number](https://docx-kit.ntnyq.dev/plugins/page-number)
- [Property Table](https://docx-kit.ntnyq.dev/plugins/property-table)
- [QR Code](https://docx-kit.ntnyq.dev/plugins/qrcode)
- [Signature Block](https://docx-kit.ntnyq.dev/plugins/signature-block)
- [Timeline](https://docx-kit.ntnyq.dev/plugins/timeline)
- [TOC](https://docx-kit.ntnyq.dev/plugins/toc)
- [Watermark](https://docx-kit.ntnyq.dev/plugins/watermark)

→ [Plugin Documentation](https://docx-kit.ntnyq.dev/plugins/)

## Contributing and Security

See [CONTRIBUTING.md](./CONTRIBUTING.md) for the toolchain, test matrix, visual
regression, and performance workflow. Report suspected vulnerabilities
privately by following [SECURITY.md](./SECURITY.md).

## Credits

This project is built on [docx](https://github.com/dolanmiu/docx).

## License

[MIT](./LICENSE) © 2025-PRESENT [ntnyq](https://github.com/ntnyq)
