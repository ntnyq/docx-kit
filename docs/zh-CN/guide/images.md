# 图片

在文档中嵌入图片，支持尺寸、浮动布局与替代文本。

## 基础图片

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

## 图片数据来源

图片可以使用多种数据格式：

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

## 本地图片路径

Node.js 入口还支持本地路径和 `file:` URL：

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx().image({ data: './logo.png', width: 200 })
```

导出时解析路径，相对路径基于进程工作目录。核心或浏览器入口不会自动读取本地文件或获取图片 URL。请提供字节或 data URL，或配置 `resolveImage(source)`，从应用控制的来源返回字节或 Blob。

## 尺寸

图片可以使用绝对尺寸。省略一个维度，
即可保持原图宽高比：

```ts
// Both dimensions specified
doc.image({ data: imageBytes, width: 400, height: 300 })

// Width only (height auto-calculates)
doc.image({ data: imageBytes, width: 200 })

// Height only
doc.image({ data: imageBytes, height: 150 })

```

## 图片格式

根据数据自动检测，也可以显式指定：

```ts
doc.image({
  data: imageBytes,
  width: 200,
  height: 200,
  imageType: 'png', // 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp'
})
```

## 浮动布局

设置图片浮动和文字环绕：

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

## 替代文本

添加替代文本以改善无障碍访问：

```ts
doc.image({
  data: chartBytes,
  width: 400,
  height: 300,
  alt: 'Quarterly revenue bar chart showing growth from Q1 to Q4',
})
```

## 段落内图片

图片也可以作为行内内容（段落的子节点）：

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

## 浏览器：从输入框加载图片

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

## 浏览器：从 URL 加载图片

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

> **注意：** 从 URL 加载图片需要服务端提供 CORS 响应头，或通过代理获取。

## 完整图片示例

### 带图片的产品目录

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

### 插件生成的图表

ECharts 等插件自动将图表渲染为图片，详见[插件指南](/zh-CN/guide/plugins)。
