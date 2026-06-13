# Getting Started

## Installation

```bash
pnpm add docx-kit
```

`docx-kit` depends on [`docx`](https://github.com/dolanmiu/docx) as its runtime engine — no additional setup needed.

### Optional Peer Dependencies

Some built-in plugins require additional packages:

```bash
# For QRCode plugin
pnpm add qrcode

# For ECharts plugin (browser only)
pnpm add echarts
```

## Quick Start

Create your first .docx document in 3 lines:

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx()

await doc
  .h1('Hello, World')
  .p('This is my first docx-kit document.')
  .save('hello.docx')
```

## Core Concepts

docx-kit is built around three core concepts:

| Concept     | Description                                                                 |
| ----------- | --------------------------------------------------------------------------- |
| **Styles**  | CSS-like style rules (fontSize, color, margin...) organized as a stylesheet |
| **Nodes**   | Content building blocks: headings, paragraphs, tables, images, plugins      |
| **Builder** | Fluent API that chains nodes together and exports to .docx                  |

### 1. Define Styles (CSS-like)

```ts
import { defineStyles } from 'docx-kit'

const styles = defineStyles({
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#1a1a2e',
    textAlign: 'center',
  },
  subtitle: {
    fontSize: 16,
    color: '#666',
    textAlign: 'center',
  },
  body: {
    fontSize: 12,
    lineHeight: 1.5,
  },
  highlight: {
    backgroundColor: '#fff3cd',
    color: '#856404',
  },
})
```

### 2. Build Content (Fluent API)

```ts
const doc = createDocx({ styles })

await doc
  .h1('Quarterly Report', { className: 'title' })
  .p('Q1 2026 Financial Summary', { className: 'subtitle' })
  .p('Revenue grew 15% year-over-year...', { className: 'body' })
  .p('⚠️ Note: exchange rate fluctuations', { className: 'highlight' })
  .save('report.docx')
```

### 3. Export

```ts
// Save to disk (Node.js)
await doc.save('output.docx')

// Get as Blob (browser)
const blob = await doc.toBlob()

// Get as Uint8Array (cross-platform)
const bytes = await doc.toBuffer()

// Get as base64 string
const b64 = await doc.toBase64()
```

## Configuration

Full configuration with page setup, metadata, and element defaults:

```ts
const doc = createDocx({
  // Page setup
  page: {
    size: 'A4',
    orientation: 'portrait',
    margin: '20mm 25mm',
  },

  // Named styles
  styles: defineStyles({
    title: { fontSize: 24, fontWeight: 'bold' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),

  // Document metadata (appears in File → Info)
  metadata: {
    title: 'Annual Report 2026',
    creator: 'docx-kit',
    subject: 'Financial Report',
    description: 'Q1–Q4 financial performance summary',
    keywords: ['report', 'finance', '2026'],
  },

  // Element-level defaults
  defaults: {
    text: { fontFamily: 'Calibri', fontSize: 11 },
    paragraph: { textAlign: 'left' },
    table: { bordered: true },
    cell: { verticalAlign: 'middle' },
  },
})
```

## What's Next?

You now know enough to build any document. Choose your next step based on what you need:

### I want to start from a pre-built style

→ [Style Presets](/guide/presets) — `classic`, `modern`, `academic` ready-to-use configs

```ts
import { createDocx, modernPreset } from 'docx-kit'
const doc = createDocx(modernPreset.config)
```

### I want a consistent look across many documents

→ [Themes](/guide/themes) — `minimal`, `ocean`, `warm` with token-based color/font systems

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'
const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({ title: { color: '$colors.primary' } }),
})
```

### I want AI/LLM to generate my document

→ [AI Templates](/guide/ai-templates) — 4 built-in templates + prompt builder

```ts
import { reportTemplate } from 'docx-kit/ai'
const schema = reportTemplate.generate({ title: 'Q2 Report' /* ... */ })
```

### I want to integrate with Claude / ChatGPT via MCP

→ [MCP Server](/guide/mcp-server) — 6 tools + 1 resource for AI agents

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
// expose to MCP-compatible clients
```

### I want to learn the full builder API

→ [Builder API](/guide/builder-api) — Every method, every option

### I want to use a built-in plugin (QR, charts, signatures...)

→ [Plugins guide](/guide/plugins) — 12 built-in plugins

## Next Steps (Detailed)

- [Builder API](/guide/builder-api) — Learn all fluent API methods
- [CSS-like Styling](/guide/styling) — Master the styling system
- [Tables](/guide/tables) — Create data tables
- [Plugins](/guide/plugins) — Extend with custom plugins
