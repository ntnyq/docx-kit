# Node.js 與瀏覽器

docx-kit 同時支援 Node.js 與瀏覽器環境。平臺專屬 API 分別位於 `docx-kit/node` 和 `docx-kit/browser`。

## 平臺支援矩陣

| API | 跨平臺（`docx-kit`） | Node.js（`docx-kit/node`） | 瀏覽器（`docx-kit/browser`） |
|---|---|---|---|
| `createDocx()` | ✅ | — | — |
| `renderDocx()` | ✅ | — | — |
| `defineStyles()` | ✅ | — | — |
| `definePlugin()` | ✅ | — | — |
| `DocxBuilder` | ✅ | — | — |
| `barcodePlugin` | ✅ | ✅ | ✅ |
| `qrcodePlugin` | ✅ | — | — |
| `echartsPlugin` | ✅（瀏覽器） | ❌ | ✅（內建） |
| `dataUrlToUint8Array()` | ✅ | ✅ | ✅ |
| `saveDocument()` | — | ✅ | ❌ |
| `normalizeImageData()` | — | ❌ | ✅ |
| `DocxKitError` | ✅ | — | — |

## 跨平臺（預設入口）

從 `'docx-kit'` 匯入可在各平臺使用的建構器 API：

```ts
import { createDocx, defineStyles } from 'docx-kit'

const doc = createDocx({
  styles: defineStyles({ p: { fontSize: 12 } }),
})

doc
  .h1('Hello')
  .p('This works in both Node.js and browser.')
```

## Node.js 平臺

從 `'docx-kit/node'` 匯入 Node.js 專屬 API：

```ts
import { createDocx } from 'docx-kit'
import { saveDocument, dataUrlToUint8Array } from 'docx-kit/node'
```

### Node.js 可用 API

#### 流式輸出

`save()` 將編譯後的文件直接流式寫入磁盤。當其他 Node.js 目標負責寫入時，
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

使用 `Buffer` 解碼 base64 data URL：

```ts
import { dataUrlToUint8Array } from 'docx-kit/node'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

### Node.js 不可用 API

| API | 原因 | 替代方案 |
|---|---|---|
| `echartsPlugin` | 需要瀏覽器 `window` 或 DOM | 使用伺服器端 canvas（如 `node-canvas` + `echarts`），或在用戶端預呈現圖表 |
| `normalizeImageData()` | Node.js 中較少用 `Blob` 處理圖片 | 使用 `new Uint8Array(await blob.arrayBuffer())` |

## 瀏覽器平臺

從 `'docx-kit/browser'` 匯入瀏覽器專屬 API：

```ts
import { createDocx, echartsPlugin } from 'docx-kit'
import { dataUrlToUint8Array, normalizeImageData } from 'docx-kit/browser'
```

### 瀏覽器可用 API

#### `dataUrlToUint8Array(dataUrl)`

使用 `atob()` 解碼 base64 data URL：

```ts
import { dataUrlToUint8Array } from 'docx-kit/browser'

const bytes = await dataUrlToUint8Array('data:image/png;base64,iVBORw0KGgo...')
```

#### `normalizeImageData(data)`

將 `Blob` 轉換為 `Uint8Array`：

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

### 瀏覽器不可用 API

| API | 原因 | 替代方案 |
|---|---|---|
| `saveDocument()` | 瀏覽器沒有檔案系統訪問能力 | 使用 `doc.toBlob()` + `URL.createObjectURL()` 觸發下載 |

### 瀏覽器下載模式

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

內建 ECharts 外掛需要瀏覽器 DOM。伺服器端呈現圖表時，請提供自訂呈現函式：

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

## 跨平臺 dataUrlToUint8Array

跨平臺版本會自動檢測環境：

```ts
import { dataUrlToUint8Array } from 'docx-kit'

// Works in Node.js (Buffer) and browser (atob)
const bytes = await dataUrlToUint8Array('data:image/png;base64,...')
```

內部檢查 `typeof atob === 'function'`：
- 存在 `atob` → 使用瀏覽器路徑
- 否則 → 動態 `import('node:buffer')`
