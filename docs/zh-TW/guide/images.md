# 圖片

在文件中嵌入圖片，支援尺寸、浮動佈局與替代文字。

## 基本圖片

```ts
import { createDocx } from 'docx-kit/node'
import { readFileSync } from 'node:fs'

const imageBytes = readFileSync('./logo.png')

const doc = createDocx()

doc
  .h1('Report')
  .image({
    data: imageBytes,
    width: 200,
    height: 80,
    alt: 'Company Logo',
  })
  .p('Report content...')
  .save('with-logo.docx')
```

## 圖片資料來源

圖片可以使用多種資料格式：

```ts
// From file (Node.js)
import { readFileSync } from 'node:fs'
doc.image({ data: readFileSync('./chart.png'), width: 400, height: 300 })

// From Uint8Array
const bytes = new Uint8Array([...])
doc.image({ data: bytes, width: 200, height: 200 })

// From ArrayBuffer
const buffer = await fetch('/image.png').then(r => r.arrayBuffer())
doc.image({ data: buffer, width: 200, height: 200 })

// From Blob (browser)
const blob = new Blob([...], { type: 'image/png' })
doc.image({ data: blob, width: 200, height: 200 })

// From base64 data URL
const dataUrl = 'data:image/png;base64,iVBORw0K...'
doc.image({ data: dataUrl, width: 200, height: 200 })
```

## 本地圖片路徑

Node.js 入口還支援本地路徑和 `file:` URL：

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx().image({ data: './logo.png', width: 200 })
```

匯出時解析路徑，相對路徑基於進程工作目錄。核心或瀏覽器入口不會自動讀取本地檔案或獲取圖片 URL。請提供位元組或 data URL，或設定 `resolveImage(source)`，從應用控制的來源回傳位元組或 Blob。

## 尺寸

圖片可以使用絕對尺寸。省略一個維度，
即可保持原圖寬高比：

```ts
// Both dimensions specified
doc.image({ data: imageBytes, width: 400, height: 300 })

// Width only (height auto-calculates)
doc.image({ data: imageBytes, width: 200 })

// Height only
doc.image({ data: imageBytes, height: 150 })

```

## 圖片格式

根據資料自動檢測，也可以顯式指定：

```ts
doc.image({
  data: imageBytes,
  width: 200,
  height: 200,
  imageType: 'png', // 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp'
})
```

## 浮動佈局

設置圖片浮動和文字環繞：

```ts
// Simple floating
doc.image({
  data: imageBytes,
  width: 150,
  height: 150,
  floating: true,
})

// With wrap mode and position
doc.image({
  data: imageBytes,
  width: 150,
  height: 150,
  floating: {
    wrap: 'square',  // 'square' | 'tight' | 'topAndBottom'
    x: 100,          // horizontal offset
    y: 50,           // vertical offset
  },
})
```

## 替代文字

添加替代文字以改善無障礙訪問：

```ts
doc.image({
  data: chartBytes,
  width: 400,
  height: 300,
  alt: 'Quarterly revenue bar chart showing growth from Q1 to Q4',
})
```

## 段落內圖片

圖片也可以作為行內內容（段落的子節點）：

```ts
doc
  .p('Here is an inline icon: ')
  .add({
    type: 'paragraph',
    children: [
      { type: 'text', text: 'Click here → ' },
      { type: 'image', data: iconBytes, width: 16, height: 16 },
      { type: 'text', text: ' to continue.' },
    ],
  })
```

## 瀏覽器：從輸入框載入圖片

```ts
// HTML: <input type="file" id="fileInput" accept="image/*" />

import { createDocx, normalizeImageData } from 'docx-kit/browser'

const input = document.querySelector('#fileInput') as HTMLInputElement
const file = input.files![0]

// Browser Blob → normalized data for ImageRun
const data = await normalizeImageData(file)

const doc = createDocx()
doc
  .h1('User-Uploaded Image')
  .image({ data, width: 300, height: 200 })

const url = URL.createObjectURL(await doc.toBlob())
const anchor = document.createElement('a')
anchor.href = url
anchor.download = 'upload.docx'
anchor.click()
URL.revokeObjectURL(url)
```

## 瀏覽器：從 URL 載入圖片

```ts
import { createDocx } from 'docx-kit'

async function loadImage(url: string) {
  const response = await fetch(url)
  const blob = await response.blob()
  return blob
}

const doc = createDocx()
const imageBlob = await loadImage('https://example.com/chart.png')

doc
  .h1('Remote Image')
  .image({ data: imageBlob, width: 500, height: 350 })
```

> **注意：** 從 URL 載入圖片需要伺服器端提供 CORS 響應頭，或透過代理獲取。

## 完整圖片範例

### 帶圖片的產品目錄

```ts
import { createDocx } from 'docx-kit/node'
import { readFileSync } from 'node:fs'

const doc = createDocx({
  page: { size: 'A4', margin: '15mm' },
})

doc
  .h1('Product Catalog')
  .h2('Widget Pro')
  .image({
    data: readFileSync('./widget-pro.png'),
    width: 300,
    height: 200,
    alt: 'Widget Pro product image',
  })
  .p('The Widget Pro is our flagship product...', {
    style: { lineHeight: 1.5 },
  })
  .pageBreak()
  .h2('Gadget Max')
  .image({
    data: readFileSync('./gadget-max.png'),
    width: 300,
    height: 200,
    alt: 'Gadget Max product image',
  })
  .p('The Gadget Max features...', {
    style: { lineHeight: 1.5 },
  })
  .save('catalog.docx')
```

### 外掛產生的圖表

ECharts 等外掛自動將圖表呈現為圖片，詳見[外掛指南](/zh-TW/guide/plugins)。
