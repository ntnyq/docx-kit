# Plugins

docx-kit ships with **12 built-in plugins** that cover common document-generation needs. Each plugin can be registered via `.use()` and invoked with `.plugin(name, options)`.

## Plugin Overview

| Plugin | Node Name | Description |
|--------|-----------|-------------|
| [Callout](./callout) | `callout` | Colored info / warning / success / danger boxes |
| [Code Block](./code-block) | `codeBlock` | Syntax-highlighted code blocks with line numbers |
| [Cover Page](./cover-page) | `coverPage` | Professional title page with title, author, date |
| [Data Table](./data-table) | `dataTable` | Auto-inferred table from an array of objects |
| [ECharts](./echarts) | `echarts` | ECharts charts rendered as embedded images |
| [Meeting Minutes](./meeting-minutes) | `meetingMinutes` | Structured meeting notes with agenda table |
| [Page Number](./page-number) | `pageNumber` | Page number field for headers/footers |
| [Property Table](./property-table) | `propertyTable` | Key-value pair table with styled cells |
| [QR Code](./qrcode) | `qrcode` | QR code images from text or URLs |
| [Signature Block](./signature-block) | `signatureBlock` | Signature lines for contracts and approvals |
| [Timeline](./timeline) | `timeline` | Chronological timeline as a styled table |
| [Watermark](./watermark) | `watermark` | Text watermark for document branding |

## Usage Pattern

Every plugin follows the same pattern:

```ts
import { createDocx, somePlugin } from 'docx-kit'

const doc = createDocx()
  .use(somePlugin())           // 1. Register the plugin
  .plugin('someName', {        // 2. Invoke with options
    // ... plugin-specific options
  })
```

## Custom Plugins

You can also create your own plugins with `definePlugin()`. See the [Plugins Guide](/guide/plugins#custom-plugins) for details.

## Platform Notes

- **ECharts** requires a browser DOM environment. In Node.js, provide a server-side canvas implementation.
- **QR Code** requires the `qrcode` peer dependency (`pnpm add qrcode`).
- **Code Block** syntax highlighting requires the optional `highlight.js` peer dependency.
- All other plugins work in both Node.js and browser with no extra dependencies.
