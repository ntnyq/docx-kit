# DOCX 预览示例

完整端到端示例：**创建文档、导出为 Blob，
并在浏览器预览**，全过程均在浏览器运行，无需服务端。

---

## 基础预览

```ts
import { createDocx, defineStyles } from 'docx-kit'
import { createDocxPreview } from '@docxkit/renderer'

// 1. Create a document
const doc = createDocx({
  styles: defineStyles({
    title: { fontSize: 24, fontWeight: 'bold', textAlign: 'center' },
    body:  { fontSize: 12, lineHeight: 1.5 },
  }),
})

doc
  .h1('Quarterly Report — Q2 2026', { className: 'title' })
  .p('This report summarizes our Q2 performance across all departments.', {
    className: 'body',
  })
  .table({
    columns: [
      { key: 'metric', title: 'Metric' },
      { key: 'q1', title: 'Q1' },
      { key: 'q2', title: 'Q2' },
      { key: 'change', title: 'Change' },
    ],
    data: [
      { metric: 'Revenue', q1: '$1.2M', q2: '$1.5M', change: '+25%' },
      { metric: 'Users', q1: '8,400', q2: '10,200', change: '+21%' },
    ],
  })

// 2. Export to Blob
const blob = await doc.toBlob()

// 3. Preview in browser
const preview = createDocxPreview(document.getElementById('preview')!)
await preview.render(blob)
```

---

## 上传文件预览

让用户上传 `.docx` 文件并立即预览：

```html
<input type="file" id="file-input" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
<div id="preview"></div>
```

```ts
import { createDocxPreview } from '@docxkit/renderer'

const input = document.getElementById('file-input') as HTMLInputElement
const preview = createDocxPreview(document.getElementById('preview')!)

input.addEventListener('change', async () => {
  const file = input.files?.[0]
  if (!file) return

  // Validate extension
  if (!file.name.endsWith('.docx')) {
    alert('Please upload a .docx file')
    return
  }

  await preview.render(file)
})
```

---

## 远程 URL 预览

从服务端获取并预览文档：

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(container, {
  pageMode: 'continuous',  // single scroll
  renderHeaders: true,
  renderFooters: true,
})

// URL input — fetched automatically
await preview.render('https://example.com/reports/2026-Q2.docx')
```

> **CORS 提示：** 托管 `.docx` 文件的服务端必须允许跨域请求，
> 否则获取文件会失败。请使用代理，或在服务端配置 CORS 响应头。

---

## 预览控制（分页）

构建简单的上一页、下一页控件：

```html
<div id="preview"></div>
<button id="prev">Previous</button>
<button id="next">Next</button>
<span id="page-info"></span>
```

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(document.getElementById('preview')!, {
  pageMode: 'paged',
})

await preview.render(blob)

// docx-preview doesn't expose a page API directly.
// The rendered pages are DOM elements with class 'docx-page'.
// You can implement custom pagination by manipulating scroll position:

function getCurrentPage(): number {
  const container = preview.container
  const pages = container.querySelectorAll('.docx-page')
  const scrollTop = container.scrollTop
  for (let i = 0; i < pages.length; i++) {
    if (pages[i]!.getBoundingClientRect().top >= 0) return i
  }
  return pages.length - 1
}
```

---

## Microsoft Office Online 回退方案

文档有公开 URL 时，可使用此方案获得更高还原度：

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(container, {
  renderer: 'microsoft',
  // Optional: use a self-hosted Office Online Server
  // microsoftViewerUrl: 'https://oos.example.com/wv/embed?src=',
})

await preview.render('https://example.com/public/report.docx')
```

> 文档 URL 必须可公开访问，微软的服务器会获取该文件。

---

## 完整 HTML 页面示例

包含上传与预览功能的独立 HTML 页面：

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>DOCX Preview</title>
  <style>
    body { font-family: sans-serif; margin: 2rem; }
    #toolbar { margin-bottom: 1rem; }
    #preview {
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      min-height: 600px;
      padding: 1rem;
    }
  </style>
</head>
<body>
  <h1>DOCX Preview</h1>

  <div id="toolbar">
    <input type="file" id="file-input" accept=".docx">
    <button id="clear-btn">Clear</button>
  </div>

  <div id="preview"></div>

  <script type="module">
    import { createDocxPreview } from '@docxkit/renderer'

    const input = document.getElementById('file-input')
    const clearBtn = document.getElementById('clear-btn')
    const preview = createDocxPreview(document.getElementById('preview'), {
      className: 'my-preview',
      pageMode: 'paged',
    })

    input.addEventListener('change', async () => {
      const file = input.files?.[0]
      if (!file) return
      await preview.render(file)
    })

    clearBtn.addEventListener('click', () => {
      preview.clear()
      input.value = ''
    })

    // Clean up on page unload
    window.addEventListener('beforeunload', () => {
      preview.destroy()
    })
  </script>
</body>
</html>
```

---

## 错误处理

```ts
import { createDocxPreview, PREVIEW_ERROR_CODES } from '@docxkit/renderer'
import { DocxKitError } from '@docxkit/core'

const preview = createDocxPreview(container)

try {
  await preview.render(invalidInput)
} catch (err) {
  if (err instanceof DocxKitError) {
    switch (err.code) {
      case PREVIEW_ERROR_CODES.PREVIEW_INPUT_INVALID:
        showError('Please upload a valid .docx file.')
        break
      case PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED:
        showError('Failed to download the document. Check the URL and CORS settings.')
        break
      case PREVIEW_ERROR_CODES.PREVIEW_RENDER_FAILED:
        showError('This document could not be rendered. It may be corrupt or use unsupported features.')
        break
    }
  }
}
```

---

## 相关文档

- [浏览器预览指南](/zh-CN/guide/preview) — 完整 API 参考
- [Builder API](/zh-CN/guide/builder-api) — 生成待预览的文档
- [运行平台](/zh-CN/guide/platforms) — 浏览器 `toBlob()` 与 Node.js `save()`
