# DOCX 預覽範例

完整端到端範例：**建立文件、匯出為 Blob，
並在瀏覽器預覽**，全過程均在瀏覽器執行，無需伺服器端。

---

## 基本預覽

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

## 上傳檔案預覽

讓用戶上傳 `.docx` 檔案並立即預覽：

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

## 遠程 URL 預覽

從伺服器端獲取並預覽文件：

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

> **CORS 提示：** 託管 `.docx` 檔案的伺服器端必須允許跨域請求，
> 否則獲取檔案會失敗。請使用代理，或在伺服器端設定 CORS 響應頭。

---

## 預覽控制（分頁）

建置簡單的上一頁、下一頁控制項：

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

## Microsoft Office Online 備援方案

文件有公開 URL 時，可使用此方案獲得更高還原度：

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(container, {
  renderer: 'microsoft',
  // Optional: use a self-hosted Office Online Server
  // microsoftViewerUrl: 'https://oos.example.com/wv/embed?src=',
})

await preview.render('https://example.com/public/report.docx')
```

> 文件 URL 必須可公開訪問，微軟的伺服器會獲取該檔案。

---

## 完整 HTML 頁面範例

包含上傳與預覽功能的獨立 HTML 頁面：

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

## 錯誤處理

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

## 相關文件

- [瀏覽器預覽指南](/zh-TW/guide/preview) — 完整 API 參考
- [Builder API](/zh-TW/guide/builder-api) — 產生待預覽的文件
- [執行平臺](/zh-TW/guide/platforms) — 瀏覽器 `toBlob()` 與 Node.js `save()`
