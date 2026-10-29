# 瀏覽器預覽

**直接在瀏覽器中**呈現 `.docx` 檔案，無需伺服器端、上傳檔案或轉換步驟。

`@docxkit/renderer` 透過
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs) 將 Word 文件呈現到網頁，採用基於 DOM 的高還原度佈局，也可備援到 Microsoft Office Online iframe。

---

## 安裝

```sh
npm install @docxkit/renderer
```

> 自動安裝 `docx-preview`（Apache-2.0，gzip 後約 47 kB）與 `jszip`。

---

## 快速開始

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(
  document.getElementById('preview')!,
)

// Render a Blob (e.g. from docx-kit's toBlob())
await preview.render(blob)

// Later — clean up
preview.destroy()
```

---

## 輸入型別

`preview.render(input)` 自動檢測輸入型別：

| 型別 | 範例 | 說明 |
|---|---|---|
| `Blob` | `await doc.toBlob()` | 最常見，來自 `DocxBuilder.toBlob()` |
| `File` | `fileInput.files[0]!` | 來自 `<input type="file">` |
| `ArrayBuffer` | `await blob.arrayBuffer()` | 來自 `fetch()` 響應 |
| `Uint8Array` | `new Uint8Array(buffer)` | 二進制資料來源 |
| `string`（URL） | `'https://example.com/doc.docx'` | 透過 `fetch()` 獲取 |

```ts
// Blob — from docx-kit
const doc = createDocx(/* ... */)
const blob = await doc.toBlob()
await preview.render(blob)

// File — from <input type="file">
const file = document.querySelector('input[type="file"]')!.files[0]!
await preview.render(file)

// URL — fetch from a server
await preview.render('https://example.com/report.docx')
```

---

## 呈現器模式

### DOM 呈現器（預設）

使用 `docx-preview`，將 DOCX 呈現為容器內帶樣式的 HTML 元素。

預設禁用嵌入 HTML（`altChunk`）。設置 `renderAltChunks: true` 後，可在不允許指令碼和同源權限的沙箱 iframe 中展示。`javascript:` 等可執行超連結協議會被移除。嵌入 HTML 仍可能請求外部資源，因此對禁止訪問外部伺服器的不受信任文件，應保持禁用。


每次 DOM 呈現都有獨立的 CSS 作用域，嵌入字型也使用獨立名稱。產生的樣式在掛載前經過解析，定位內容受到包裹層約束。即使設定 `inWrapper: false`，該隔離包裹層仍然存在；該選項僅控制 `docx-preview` 的裝飾包裹層。原有類別名稱和 `container.querySelector()` 查詢方式繼續可用。

- **文字可選擇、可搜尋**（基於 DOM）
- 支援頁眉、頁腳、分頁、表格與圖片
- 還原度高，儘量貼近 Word 佈局

```ts
const preview = createDocxPreview(container, {
  renderer: 'dom',  // default
  pageMode: 'paged',  // or 'continuous'
})
```

### Microsoft Office Online（iframe 備援）

透過 `<iframe>` 嵌入微軟免費的在線查看器。

- **需要公開可訪問的 URL**，微軟伺服器必須能獲取檔案
- 使用微軟自己的呈現器，提供更高的 Word 兼容性
- 不支援編程控制

```ts
const preview = createDocxPreview(container, {
  renderer: 'microsoft',
})

// Only URL strings are supported
await preview.render('https://example.com/public/doc.docx')
```

對於非 URL 輸入，Microsoft 呈現器會拋出明確錯誤：

```
[MICROSOFT_URL_REQUIRED]
Microsoft renderer requires a publicly accessible URL string.
Received Blob. Upload the DOCX to a public URL first,
or use renderer: 'dom'.
```

---

## 分頁

### 分頁模式（預設）

每頁呈現為帶陰影的獨立白色區域，呈現紙張效果。

```ts
const preview = createDocxPreview(container, {
  pageMode: 'paged',  // default
})
```

### 連續模式

呈現為連續滾動的文件，不顯示分頁。

```ts
const preview = createDocxPreview(container, {
  pageMode: 'continuous',
})
```

---

## API 參考

### `createDocxPreview(container, options?)`

建立綁定到 DOM 容器的預覽實例。

**參數：**
- `container: HTMLElement` — 目標 DOM 元素
- `options: DocxPreviewOptions` — 可選設定

**回傳：** `DocxPreview`

---

### `DocxPreview` 實例

| 方法 | 說明 |
|---|---|
| `render(input)` | 呈現 DOCX 輸入（Blob、File、ArrayBuffer、Uint8Array 或 URL） |
| `clear()` | 移除呈現內容並釋放媒體 URL，保留實例 |
| `destroy()` | 完整清理：釋放 URL、清空 DOM 並標記為已銷燬 |

| 屬性 | 型別 | 說明 |
|---|---|---|
| `container` | `readonly HTMLElement` | 綁定的容器元素 |
| `currentInput` | `readonly DocxInput \| null` | 最近呈現的輸入 |

```ts
const preview = createDocxPreview(container)

// Render
await preview.render(blob)

// Re-render (replaces previous content)
await preview.render(anotherBlob)

// Clear without destroying
preview.clear()

// Full cleanup
preview.destroy()
// preview.render(...)  ← throws "has been destroyed"
```

---

## 選項

支援透傳 `docx-preview` 的全部選項：

| 選項 | 型別 | 預設值 | 說明 |
|---|---|---|---|
| `renderer` | `'dom' \| 'microsoft'` | `'dom'` | 呈現後端 |
| `pageMode` | `'paged' \| 'continuous'` | `'paged'` | 頁面顯示模式 |
| `className` | `string` | `'docx-kit-preview'` | 容器 CSS 類名 |
| `inWrapper` | `boolean` | `true` | 在頁面外添加 `<div>` 包裝 |
| `breakPages` | `boolean` | `true` | 啟用分頁，覆蓋 `pageMode` |
| `renderHeaders` | `boolean` | `false` | 呈現頁眉 |
| `renderFooters` | `boolean` | `false` | 呈現頁腳 |
| `renderFootnotes` | `boolean` | `false` | 呈現腳註 |
| `renderEndnotes` | `boolean` | `false` | 呈現尾註 |
| `ignoreFonts` | `boolean` | `false` | 不載入網頁字型 |
| `useBase64URL` | `boolean` | `false` | 使用 base64 圖片，替代 ObjectURL |
| `renderAltChunks` | `boolean` | `false` | 在受限沙箱 iframe 中呈現嵌入 HTML |
| `microsoftViewerUrl` | `string` | Office Online URL | 自訂 OOS 伺服器端 |

---

## 錯誤處理

所有錯誤都是帶結構化 `code` 的 `DocxKitError` 實例：

```ts
import { createDocxPreview, PREVIEW_ERROR_CODES } from '@docxkit/renderer'
import { DocxKitError } from '@docxkit/core'

try {
  await preview.render(invalidInput)
} catch (err) {
  if (err instanceof DocxKitError) {
    switch (err.code) {
      case PREVIEW_ERROR_CODES.PREVIEW_INPUT_INVALID:
        console.error('Invalid input type:', err.message)
        break
      case PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED:
        console.error('Network error:', err.message)
        break
      case PREVIEW_ERROR_CODES.PREVIEW_RENDER_FAILED:
        console.error('Render failed:', err.message)
        break
      case PREVIEW_ERROR_CODES.MICROSOFT_URL_REQUIRED:
        console.error('Microsoft mode needs URL:', err.message)
        break
    }
  }
}
```

| 程式碼 | 觸發條件 |
|---|---|
| `PREVIEW_INPUT_INVALID` | 輸入為 `null`、`undefined` 或無法識別的型別 |
| `PREVIEW_FETCH_FAILED` | 對 URL 的 `fetch()` 失敗（網絡或 HTTP 錯誤） |
| `PREVIEW_RENDER_FAILED` | DOCX 解析或 DOM 呈現失敗 |
| `MICROSOFT_URL_REQUIRED` | 向 Microsoft 呈現器傳入非 URL 輸入 |

---

## 限制

- **僅瀏覽器可用**：呈現器使用 DOM API，無法在 Node.js 執行
- **`docx-preview` 限制**：超大文件（200 頁以上）可能較慢；修訂和註解呈現仍為實驗功能
- **Microsoft 模式**：需要公開 URL，文件會發送到微軟伺服器
- **中日韓字型**：用戶系統安裝相應字型時可用；`docx-preview` 沿用瀏覽器字型處理

---

## 相關文件

- [範例：DOCX 預覽](/zh-TW/examples/preview) — 端到端範例
- [Builder API](/zh-TW/guide/builder-api) — 產生待預覽的文件
- [執行平臺](/zh-TW/guide/platforms) — Node.js `save()` 與瀏覽器 `toBlob()`
