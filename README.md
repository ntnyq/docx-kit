# docx-kit

[![CI](https://github.com/ntnyq/docx-kit/workflows/CI/badge.svg)](https://github.com/ntnyq/docx-kit/actions)
[![NPM VERSION](https://img.shields.io/npm/v/docx-kit.svg)](https://www.npmjs.com/package/docx-kit)
[![NPM DOWNLOADS](https://img.shields.io/npm/dy/docx-kit.svg)](https://www.npmjs.com/package/docx-kit)
[![LICENSE](https://img.shields.io/github/license/ntnyq/docx-kit.svg)](https://github.com/ntnyq/docx-kit/blob/main/LICENSE)

**CSS-like DOCX API Kit** — Type-safe, plugin-extensible Word document generation for Node.js & browser.

Built on [docx](https://github.com/dolanmiu/docx).

---

## 📖 Documentation

👉 **Full guides, API reference, plugins, and examples: https://docx-kit.ntnyq.dev**

---

## Features

- **CSS-like styling** — `fontSize`, `color`, `margin`, `border`... just like CSS
- **Fluent builder API** — `.h1()` `.p()` `.table()` `.image()` `.use().plugin()`
- **19 built-in plugins** — Barcodes, QR codes, ECharts, code blocks, invoices, TOCs, and more
- **Style presets** — `classic` (gov-doc), `modern` (business), `academic` (thesis)
- **Themes** — `minimal`, `ocean`, `warm` with token system (`$colors.primary`)
- **JSON DSL** — `renderDocx()` for AI/LLM-driven document generation
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

19 built-in plugins: Badge, Barcode, Callout, Changelog, Code Block, Cover Page, Data Table, Divider, ECharts, Invoice, Letterhead, Meeting Minutes, Page Number, Property Table, QR Code, Signature Block, Timeline, TOC, Watermark.

→ [Plugin Documentation](https://docx-kit.ntnyq.dev/plugins/)

## Contributing and Security

See [CONTRIBUTING.md](./CONTRIBUTING.md) for the toolchain, test matrix, visual
regression, and performance workflow. Report suspected vulnerabilities
privately by following [SECURITY.md](./SECURITY.md).

## License

[MIT](./LICENSE) © 2025-PRESENT [ntnyq](https://github.com/ntnyq)
