# Plugins

Plugins extend docx-kit with custom content types like QR codes, charts, signatures, and more.

## Built-in Plugins

docx-kit ships with two built-in plugins:

| Plugin | Package | Platform | Description |
|---|---|---|---|
| `qrcodePlugin()` | `qrcode` (peer dep) | Both | Generate QR code images |
| `echartsPlugin()` | `echarts` (peer dep) | Browser | Render ECharts charts as images |

## Using Plugins

### 1. Install the peer dependency

```bash
pnpm add qrcode    # for QRCode plugin
pnpm add echarts    # for ECharts plugin (browser)
```

### 2. Register and invoke

```ts
import { createDocx, qrcodePlugin } from 'docx-kit'

const doc = createDocx()
  .use(qrcodePlugin())              // register
  .h1('QR Code Demo')
  .plugin('qrcode', {               // invoke
    text: 'https://example.com',
    size: 200,
  })
  .save('qrcode.docx')
```

## QRCode Plugin

Generate QR codes from text or URLs:

```ts
import { createDocx, qrcodePlugin } from 'docx-kit'

const doc = createDocx()
  .use(qrcodePlugin())
  .h1('Business Card QR')
  .plugin('qrcode', {
    text: 'https://ntnyq.com',
    size: 256,
    errorCorrectionLevel: 'H',
    margin: 2,
    caption: 'Scan to visit my website',
  })
  .save('business-card.docx')
```

### QRCode Options

| Option | Type | Default | Description |
|---|---|---|---|
| `text` | `string` | _(required)_ | Text or URL to encode |
| `size` | `number` | `128` | QR code image size in pixels |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | `'M'` | Error correction level (~7%–30%) |
| `margin` | `number` | `1` | Quiet zone margin in modules |
| `caption` | `string` | — | Optional caption text below the QR code |

## ECharts Plugin

Render ECharts charts as embedded images (browser only):

```ts
import { createDocx, echartsPlugin } from 'docx-kit'

const doc = createDocx()
  .use(echartsPlugin())
  .h1('Sales Dashboard')
  .plugin('echarts', {
    option: {
      title: { text: 'Monthly Revenue' },
      tooltip: {},
      xAxis: {
        type: 'category',
        data: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      },
      yAxis: { type: 'value' },
      series: [
        {
          name: 'Revenue',
          type: 'bar',
          data: [120, 200, 150, 80, 70, 110],
          itemStyle: { color: '#2563eb' },
        },
      ],
    },
    width: 640,
    height: 360,
    caption: 'Figure 1: Monthly revenue by product line',
  })
  .save('dashboard.docx')
```

### Multiple Charts

```ts
doc
  .use(echartsPlugin())
  .h1('Analytics Report')

  // Bar chart
  .plugin('echarts', {
    option: {
      title: { text: 'Revenue by Quarter' },
      xAxis: { data: ['Q1', 'Q2', 'Q3', 'Q4'] },
      yAxis: {},
      series: [{ type: 'bar', data: [820, 932, 901, 1347] }],
    },
    caption: 'Figure 1: Quarterly revenue',
  })

  .pageBreak()

  // Line chart
  .plugin('echarts', {
    option: {
      title: { text: 'User Growth' },
      xAxis: { type: 'category', data: ['Jan', 'Feb', 'Mar', 'Apr', 'May'] },
      yAxis: { type: 'value' },
      series: [
        { name: 'Users', type: 'line', data: [150, 230, 224, 218, 335] },
        { name: 'Active', type: 'line', data: [120, 182, 191, 234, 290] },
      ],
    },
    caption: 'Figure 2: Monthly user growth',
  })
```

### Chart Types

Any ECharts chart type is supported — bar, line, pie, scatter, radar, heatmap, etc.:

```ts
// Pie chart
doc.plugin('echarts', {
  option: {
    title: { text: 'Market Share' },
    series: [{
      type: 'pie',
      data: [
        { name: 'Product A', value: 40 },
        { name: 'Product B', value: 30 },
        { name: 'Product C', value: 20 },
        { name: 'Product D', value: 10 },
      ],
    }],
  },
})

// Scatter chart
doc.plugin('echarts', {
  option: {
    title: { text: 'Correlation Analysis' },
    xAxis: {},
    yAxis: {},
    series: [{
      type: 'scatter',
      data: [[10, 8.04], [8, 6.95], [13, 7.58], [9, 8.81]],
    }],
  },
})
```

### ECharts Options

| Option | Type | Default | Description |
|---|---|---|---|
| `option` | `EChartsOption` | _(required)_ | Full ECharts option object |
| `width` | `number` | `640` | Chart width in pixels |
| `height` | `number` | `360` | Chart height in pixels |
| `renderer` | `'canvas' \| 'svg'` | `'canvas'` | Rendering engine |
| `imageType` | `'png' \| 'svg'` | `'png'` | Output image format |
| `caption` | `string` | — | Optional caption below the chart |

> **Node.js:** The ECharts plugin requires a browser DOM. In Node.js, it throws an error suggesting a server-side canvas solution.

## Custom Plugins

Create your own plugins with `definePlugin()`:

```ts
import { definePlugin } from 'docx-kit'
import { Paragraph, TextRun } from 'docx'

interface SignatureOptions {
  name: string
  title?: string
  date?: string
}

export function signaturePlugin() {
  return definePlugin<'signature', SignatureOptions>({
    name: 'signature',

    setup() {
      // Optional: one-time initialization
      console.log('Signature plugin loaded')
    },

    render(options) {
      const paragraphs = [
        new Paragraph({
          spacing: { after: 0 },
          children: [
            new TextRun({ text: `_________________________`, size: 22 }),
          ],
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [
            new TextRun({ text: options.name, bold: true }),
          ],
        }),
      ]

      if (options.title) {
        paragraphs.push(
          new Paragraph({
            children: [new TextRun({ text: options.title, italics: true })],
          }),
        )
      }

      if (options.date) {
        paragraphs.push(
          new Paragraph({
            children: [new TextRun({ text: `Date: ${options.date}` })],
          }),
        )
      }

      return paragraphs
    },
  })
}
```

### Using Custom Plugins

```ts
import { createDocx } from 'docx-kit'
import { signaturePlugin } from './signature-plugin'

const doc = createDocx()
  .use(signaturePlugin())
  .h1('Contract')
  .p('The parties agree to the following terms...')
  .p('Signed:')
  .plugin('signature', {
    name: 'John Doe',
    title: 'CEO',
    date: '2026-06-10',
  })
  .save('contract.docx')
```

### Plugin with Image Context

Access rendering utilities via the plugin context:

```ts
import { definePlugin, type PluginRenderContext } from 'docx-kit'
import { ImageRun, Paragraph } from 'docx'

interface BadgeOptions {
  text: string
  color: string
}

export function badgePlugin() {
  return definePlugin<'badge', BadgeOptions>({
    name: 'badge',
    async render(options, ctx: PluginRenderContext) {
      // Use ctx.utils to process images
      // Use ctx.compileNode to compile child nodes
      // Use ctx.config to access document config

      return [
        new Paragraph({
          children: [
            new TextRun({
              text: options.text,
              color: options.color,
              bold: true,
            }),
          ],
        }),
      ]
    },
  })
}
```

## Plugin System Architecture

```
┌─────────────────────────────────────────────────┐
│                  createDocx()                     │
│                      │                            │
│              DocxBuilder<TPlugins>                │
│     .use(plugin) ──► merge into TPlugins          │
│     .plugin(...) ──► add PluginNode               │
│                      │                            │
│              compileDocument()                    │
│     resolve plugin │ render(options, context)     │
│                      │                            │
│              ┌───────┴───────┐                    │
│              │   Plugin      │                    │
│              │  .render()    │                    │
│              │  returns      │                    │
│              │  Paragraph[]  │                    │
│              │  Table[]      │                    │
│              │  etc.         │                    │
│              └───────────────┘                    │
└─────────────────────────────────────────────────┘
```
