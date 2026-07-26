# Plugins

Plugins are the main extension mechanism in docx-kit. They let you add new node types without changing the core compiler, while keeping the same builder workflow and type-safe option objects.

## Mental Model

Every built-in or custom plugin follows the same lifecycle:

1. Import the plugin factory or plugin instance.
2. Register it with `.use(...)`.
3. Invoke it with `.plugin(name, options)`.

```ts
import { calloutPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'success',
    title: 'Published',
    content: 'The release notes were exported successfully.',
  })
```

## Built-in Coverage

docx-kit now includes **19 built-in plugins** across the workspace:

- Content blocks: `badge`, `callout`, `codeBlock`, `divider`, `watermark`
- Report scaffolding: `coverPage`, `letterhead`, `toc`, `pageNumber`
- Structured business data: `dataTable`, `propertyTable`, `timeline`, `invoice`, `meetingMinutes`, `changelog`
- Media and embeds: `barcode`, `echarts`, `qrcode`
- Approval workflows: `signatureBlock`

The full reference lives in the [Plugins section](/plugins/).

## Package Layout

Each built-in plugin is published as its own package under `packages-plugins/`:

```text
packages-plugins/
  badge/
  barcode/
  callout/
  changelog/
  code-block/
  cover-page/
  data-table/
  divider/
  echarts/
  invoice/
  letterhead/
  meeting-minutes/
  page-number/
  property-table/
  qrcode/
  signature-block/
  timeline/
  toc/
  watermark/
```

This keeps plugin releases, tests, peer dependencies, and documentation isolated.

## Peer Dependencies

Some plugins intentionally keep their heavyweight runtime dependencies optional:

| Plugin | Why it is optional |
| --- | --- |
| `@docxkit/plugin-barcode` | `bwip-js` is only needed when generating barcode images |
| `@docxkit/plugin-code-block` | `highlight.js` is only needed for syntax coloring |
| `@docxkit/plugin-echarts` | `echarts` is only needed when chart rendering is used |
| `@docxkit/plugin-qrcode` | `qrcode` is only needed when generating QR images |

If a plugin relies on an optional peer dependency and it is missing, the plugin either falls back gracefully or throws an actionable runtime error.

## Custom Plugins

Use `definePlugin()` when you need a new node type that is specific to your product or document workflow.

```ts
import { definePlugin } from 'docx-kit'
import { Paragraph } from 'docx'

interface DisclaimerOptions {
  text: string
}

const disclaimerPlugin = definePlugin<'disclaimer', DisclaimerOptions>({
  name: 'disclaimer',
  render(options) {
    return new Paragraph(options.text)
  },
})
```

A good custom plugin should:

- accept a small, explicit options object
- return normal `docx` nodes or arrays of nodes
- keep validation close to the render function
- avoid hiding unrelated document mutations

## Tooling Around Plugins

- `@docxkit/create-plugin` scaffolds new plugin packages
- `@docxkit/pdk` provides helpers for plugin testing
- `@docxkit/loader` loads plugins from inline, npm, URL, or local sources
- `@docxkit/registry` discovers plugins published with docx-kit keywords
- `@docxkit/mcp` exposes plugin-aware document generation tools to AI agents

See [Monorepo Structure](/ecosystem/monorepo) and [Package Catalog](/ecosystem/packages) for where each piece lives.
