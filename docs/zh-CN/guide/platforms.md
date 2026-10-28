# Node.js 与浏览器

docx-kit 同时支持 Node.js 与浏览器环境。平台专属 API 分别位于 `docx-kit/node` 和 `docx-kit/browser`。

## 平台支持矩阵

| API | 跨平台（`docx-kit`） | Node.js（`docx-kit/node`） | 浏览器（`docx-kit/browser`） |
|---|---|---|---|
| `createDocx()` | ✅ | — | — |
| `renderDocx()` | ✅ | — | — |
| `defineStyles()` | ✅ | — | — |
| `definePlugin()` | ✅ | — | — |
| `DocxBuilder` | ✅ | — | — |
| `barcodePlugin` | ✅ | ✅ | ✅ |
| `qrcodePlugin` | ✅ | — | — |
| `echartsPlugin` | ✅（浏览器） | ❌ | ✅（内置） |
| `dataUrlToUint8Array()` | ✅ | ✅ | ✅ |
| `saveDocument()` | — | ✅ | ❌ |
| `normalizeImageData()` | — | ❌ | ✅ |
| `DocxKitError` | ✅ | — | — |

## 跨平台（默认入口）

从 `'docx-kit'` 导入可在各平台使用的构建器 API：

```ts
import { createDocx, defineStyles } from 'docx-kit'

const doc = createDocx({
  styles: defineStyles({ p: { fontSize: 12 } }),
})

doc
  .h1('Hello')
  .p('This works in both Node.js and browser.')
```

## Node.js 平台

从 `'docx-kit/node'` 导入 Node.js 专属 API：

```ts
import { createDocx } from 'docx-kit'
import { saveDocument, dataUrlToUint8Array } from 'docx-kit/node'
```

### Node.js 可用 API

#### 流式输出

`save()` 将编译后的文档直接流式写入磁盘。当其他 Node.js 目标负责写入时，
可使用 `toStream()` 或 `streamDocument()`：

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

使用 `Buffer` 解码 base64 data URL：

```ts
import { dataUrlToUint8Array } from 'docx-kit/node'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

### Node.js 不可用 API

| API | 原因 | 替代方案 |
|---|---|---|
| `echartsPlugin` | 需要浏览器 `window` 或 DOM | 使用服务端 canvas（如 `node-canvas` + `echarts`），或在客户端预渲染图表 |
| `normalizeImageData()` | Node.js 中较少用 `Blob` 处理图片 | 使用 `new Uint8Array(await blob.arrayBuffer())` |

## 浏览器平台

从 `'docx-kit/browser'` 导入浏览器专属 API：

```ts
import { createDocx, echartsPlugin } from 'docx-kit'
import { dataUrlToUint8Array, normalizeImageData } from 'docx-kit/browser'
```

### 浏览器可用 API

#### `dataUrlToUint8Array(dataUrl)`

使用 `atob()` 解码 base64 data URL：

```ts
import { dataUrlToUint8Array } from 'docx-kit/browser'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

#### `normalizeImageData(data)`

将 `Blob` 转换为 `Uint8Array`：

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

### 浏览器不可用 API

| API | 原因 | 替代方案 |
|---|---|---|
| `saveDocument()` | 浏览器没有文件系统访问能力 | 使用 `doc.toBlob()` + `URL.createObjectURL()` 触发下载 |

### 浏览器下载模式

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

## Node.js 中使用 ECharts

内置 ECharts 插件需要浏览器 DOM。服务端渲染图表时，请提供自定义渲染函数：

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

## 跨平台 dataUrlToUint8Array

跨平台版本会自动检测环境：

```ts
import { dataUrlToUint8Array } from 'docx-kit'

// Works in Node.js (Buffer) and browser (atob)
const bytes = await dataUrlToUint8Array('data:image/png;base64,...')
```

内部检查 `typeof atob === 'function'`：
- 存在 `atob` → 使用浏览器路径
- 否则 → 动态 `import('node:buffer')`
