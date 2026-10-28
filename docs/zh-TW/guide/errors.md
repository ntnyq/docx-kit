# 錯誤處理

docx-kit 拋出結構化的 `DocxKitError` 實例，攜帶機器可讀的錯誤代碼。

## DocxKitError

```ts
class DocxKitError extends Error {
  readonly code: string | ErrorCode  // Machine-readable code
  readonly cause: unknown            // Underlying error (if any)
  name: 'DocxKitError'               // Always "DocxKitError"
}
```

### 捕獲錯誤

```ts
import { DocxKitError, ERROR_CODES } from 'docx-kit/node'

try {
  await doc.save('output.docx')
} catch (err) {
  if (err instanceof DocxKitError) {
    console.error(`[${err.code}] ${err.message}`)

    switch (err.code) {
      case ERROR_CODES.EXPORT_FAILED:
        console.error('Document export failed. Check disk space and permissions.')
        break
      case ERROR_CODES.PLUGIN_NOT_REGISTERED:
        console.error('Plugin not registered. Did you forget .use()?')
        break
      case ERROR_CODES.STYLE_UNKNOWN_CLASS:
        console.error('Unknown className. Check your stylesheet keys.')
        break
      default:
        console.error('Unknown docx-kit error:', err.message)
    }
  } else {
    // Non-docx-kit error
    console.error('Unexpected error:', err)
  }
}
```

## 錯誤代碼

| 常量 | 程式碼 | 說明 |
|---|---|---|
| `ERROR_CODES.EXPORT_FAILED` | `'EXPORT_FAILED'` | 文件匯出（儲存或封裝）失敗 |
| `ERROR_CODES.IMAGE_INVALID_DATA` | `'IMAGE_INVALID_DATA'` | 圖片資料為空、損壞或不受支援 |
| `ERROR_CODES.IMAGE_UNKNOWN_TYPE` | `'IMAGE_UNKNOWN_TYPE'` | 無法確定圖片格式 |
| `ERROR_CODES.PLUGIN_NOT_REGISTERED` | `'PLUGIN_NOT_REGISTERED'` | 外掛節點引用了未註冊的外掛 |
| `ERROR_CODES.PLUGIN_RENDER_FAILED` | `'PLUGIN_RENDER_FAILED'` | 外掛的 `render()` 拋出錯誤 |
| `ERROR_CODES.STYLE_INVALID_UNIT` | `'STYLE_INVALID_UNIT'` | 樣式值中的單位字串無效 |
| `ERROR_CODES.STYLE_UNKNOWN_CLASS` | `'STYLE_UNKNOWN_CLASS'` | `className` 引用了不存在的樣式表鍵 |
| `ERROR_CODES.TABLE_INVALID_COLUMNS` | `'TABLE_INVALID_COLUMNS'` | 建立表格時未提供列 |
| `ERROR_CODES.UNKNOWN_NODE_TYPE` | `'UNKNOWN_NODE_TYPE'` | 節點型別沒有已註冊的編譯器 |

## 常見錯誤場景

### 匯出失敗

```ts
try {
  await doc.save('/read-only-dir/output.docx')
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.EXPORT_FAILED) {
    // Permission denied, disk full, etc.
    // err.cause contains the original fs error
    console.error('Export failed:', err.cause)
  }
}
```

### 外掛未註冊

```ts
try {
  await doc
    // Forgot .use(qrcodePlugin())!
    .plugin('qrcode', { text: 'https://example.com' })
    .save('output.docx')
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.PLUGIN_NOT_REGISTERED) {
    console.error('Plugin "qrcode" was not registered. Call .use(qrcodePlugin()) first.')
  }
}
```

### 未知類名

```ts
const styles = defineStyles({ body: { fontSize: 12 } })

try {
  await createDocx({ styles })
    .p('Hello', { className: 'boddy' })  // typo!
    .save('output.docx')
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.STYLE_UNKNOWN_CLASS) {
    console.error('Unknown class "boddy". Did you mean "body"?')
  }
}
```

### 無效圖片資料

```ts
try {
  doc.image({ data: '' })  // empty data
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.IMAGE_INVALID_DATA) {
    console.error('Image data is empty.')
  }
}
```

### 外掛呈現失敗

```ts
// Custom plugin that might fail
const fragilePlugin = definePlugin<'fragile', { url: string }>({
  name: 'fragile',
  async render(options) {
    const response = await fetch(options.url)
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    // ...
  },
})

try {
  await createDocx()
    .use(fragilePlugin)
    .plugin('fragile', { url: 'https://invalid.example' })
    .save('output.docx')
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.PLUGIN_RENDER_FAILED) {
    // err.cause contains the original Error from fetch
    console.error('Plugin render failed:', err.message)
  }
}
```

## 最佳實踐

### 型別安全的錯誤匹配

```ts
import { DocxKitError, ERROR_CODES } from 'docx-kit/node'

function isDocxError(err: unknown, code: keyof typeof ERROR_CODES): err is DocxKitError {
  return err instanceof DocxKitError && err.code === ERROR_CODES[code]
}

try {
  await doc.save('output.docx')
} catch (err) {
  if (isDocxError(err, 'EXPORT_FAILED')) {
    // err is typed as DocxKitError with EXPORT_FAILED code
  }
}
```

### 非同步錯誤處理

所有匯出方法（`save()`、`toBlob()`、`toBuffer()`、`toBase64()`）都是非同步的，請始終使用 `await`：

```ts
// ✅ Correct
await doc.save('output.docx')

// ❌ Wrong — save() returns a Promise, not void
doc.save('output.docx')
```

### 封裝批量產生

```ts
async function generateReports(items: Item[]) {
  const results: { item: Item; error?: string }[] = []

  for (const item of items) {
    try {
      const doc = createDocx()
        .h1(item.title)
        .p(item.content)

      await doc.save(`${item.id}.docx`)
      results.push({ item })
    } catch (err) {
      const message = err instanceof DocxKitError
        ? `[${err.code}] ${err.message}`
        : String(err)
      results.push({ item, error: message })
    }
  }

  return results
}
```
