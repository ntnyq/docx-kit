# Node.js & Browser

docx-kit works in both Node.js and browser environments. Platform-specific APIs are separated under `docx-kit/node` and `docx-kit/browser`.

## Platform Matrix

| API | Cross-Platform (`docx-kit`) | Node.js (`docx-kit/node`) | Browser (`docx-kit/browser`) |
|---|---|---|---|
| `createDocx()` | ✅ | — | — |
| `renderDocx()` | ✅ | — | — |
| `defineStyles()` | ✅ | — | — |
| `definePlugin()` | ✅ | — | — |
| `DocxBuilder` | ✅ | — | — |
| `barcodePlugin` | ✅ | ✅ | ✅ |
| `qrcodePlugin` | ✅ | — | — |
| `echartsPlugin` | ✅ (browser) | ❌ | ✅ (built-in) |
| `dataUrlToUint8Array()` | ✅ | ✅ | ✅ |
| `saveDocument()` | — | ✅ | ❌ |
| `normalizeImageData()` | — | ❌ | ✅ |
| `DocxKitError` | ✅ | — | — |

## Cross-Platform (default)

Import from `'docx-kit'` for the builder API that works everywhere:

```ts
import { createDocx, defineStyles } from 'docx-kit'

const doc = createDocx({
  styles: defineStyles({ p: { fontSize: 12 } }),
})

doc
  .h1('Hello')
  .p('This works in both Node.js and browser.')
```

## Node.js Platform

Import platform-specific Node.js APIs from `'docx-kit/node'`:

```ts
import { createDocx } from 'docx-kit'
import { saveDocument, dataUrlToUint8Array } from 'docx-kit/node'
```

### Available in Node.js

#### Streaming output

`save()` streams a compiled document directly to disk. Use `toStream()` or
`streamDocument()` when another Node.js destination owns the writable side:

```ts
import { createDocx, saveDocument, streamDocument } from 'docx-kit/node'

const doc = createDocx()
  .h1('Hello')
  .p('World')

// Method 1: save() shortcut (uses saveDocument internally)
await doc.save('output.docx')

// Method 2: pipe the builder output elsewhere
const stream = await doc.toStream()
stream.pipe(uploadDestination)

// Method 3: work with a compiled Document
const compiled = await doc.toDocument()
await saveDocument(compiled, 'output.docx')
streamDocument(compiled).pipe(uploadDestination)
```

#### `dataUrlToUint8Array(dataUrl)`

Decode base64 data-URLs using `Buffer`:

```ts
import { dataUrlToUint8Array } from 'docx-kit/node'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

### Not Available in Node.js

| API | Reason | Workaround |
|---|---|---|
| `echartsPlugin` | Requires browser `window`/DOM | Use server-side canvas (e.g. `node-canvas` + `echarts`), or pre-render charts on the client |
| `normalizeImageData()` | Node rarely uses `Blob` for images | Use `new Uint8Array(await blob.arrayBuffer())` |

## Browser Platform

Import platform-specific browser APIs from `'docx-kit/browser'`:

```ts
import { createDocx, echartsPlugin } from 'docx-kit'
import { dataUrlToUint8Array, normalizeImageData } from 'docx-kit/browser'
```

### Available in Browser

#### `dataUrlToUint8Array(dataUrl)`

Decode base64 data-URLs using `atob()`:

```ts
import { dataUrlToUint8Array } from 'docx-kit/browser'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

#### `normalizeImageData(data)`

Convert `Blob` to `Uint8Array`:

```ts
import { normalizeImageData } from 'docx-kit/browser'

// From <input type="file">
const file = document.querySelector('input[type=file]').files![0]
const data = await normalizeImageData(file)
doc.image({ data, width: 200, height: 200 })

// From fetch
const blob = await fetch('/image.png').then(r => r.blob())
const data = await normalizeImageData(blob)
```

### Not Available in Browser

| API | Reason | Workaround |
|---|---|---|
| `saveDocument()` | No filesystem access in browsers | Use `doc.toBlob()` + `URL.createObjectURL()` to trigger download |

### Browser Download Pattern

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx()
  .h1('Report')
  .p('Generated in the browser.')
  .pageBreak()
  .h2('Data')
  .table({
    columns: [{ key: 'name', title: 'Name' }, { key: 'value', title: 'Value' }],
    data: [{ name: 'Item 1', value: '100' }],
  })

// Trigger download
const blob = await doc.toBlob()
const url = URL.createObjectURL(blob)

const a = document.createElement('a')
a.href = url
a.download = 'report.docx'
document.body.appendChild(a)
a.click()
document.body.removeChild(a)

// Clean up
URL.revokeObjectURL(url)
```

## ECharts in Node.js

The built-in ECharts plugin requires a browser DOM. For server-side chart rendering, provide your own render function:

```ts
// Node.js with node-canvas (conceptual example)
import { createDocx } from 'docx-kit/node'
import * as echarts from 'echarts'
import { createCanvas } from 'canvas'

// Pre-render chart to PNG buffer, then embed as image
function renderChartToBuffer(option: any): Buffer {
  const canvas = createCanvas(800, 400)
  const chart = echarts.init(canvas as any)
  chart.setOption(option)
  const buffer = canvas.toBuffer('image/png')
  chart.dispose()
  return buffer
}

const chartBuffer = renderChartToBuffer({
  xAxis: { data: ['A', 'B', 'C'] },
  yAxis: {},
  series: [{ type: 'bar', data: [1, 2, 3] }],
})

const doc = createDocx()
  .h1('Server-Rendered Chart')
  .image({ data: chartBuffer, width: 400, height: 200 })
  .save('server-chart.docx')
```

## Cross-Platform dataUrlToUint8Array

The cross-platform version auto-detects the environment:

```ts
import { dataUrlToUint8Array } from 'docx-kit'

// Works in Node.js (Buffer) and browser (atob)
const bytes = await dataUrlToUint8Array('data:image/png;base64,...')
```

Internally, it checks `typeof atob === 'function'`:
- If `atob` exists → browser path
- Otherwise → dynamic `import('node:buffer')`
