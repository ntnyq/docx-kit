# 错误处理

docx-kit 抛出结构化的 `DocxKitError` 实例，携带机器可读的错误代码。

## DocxKitError

```ts
class DocxKitError extends Error {
  readonly code: string | ErrorCode  // Machine-readable code
  readonly cause: unknown            // Underlying error (if any)
  readonly diagnosticCode?: DiagnosticCode // DOCXKIT_ code for known errors
  readonly fix?: string              // Repair guidance
  readonly docs?: string             // Canonical error reference URL
  name: 'DocxKitError'               // Always "DocxKitError"
}
```

### 捕获错误

```ts
import { DocxKitError, ERROR_CODES } from 'docx-kit/node'

try {
  await doc.save('output.docx')
} catch (err) {
  if (err instanceof DocxKitError) {
    console.error(`[${err.diagnosticCode ?? err.code}] ${err.message}`)

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

## 错误代码

所有诊断码和修复步骤见[错误参考](/zh-CN/errors/)。

`error.code` 保留旧值，`error.diagnosticCode` 新增 `DOCXKIT_` 前缀。已知错误还提供 `fix` 和 `docs`，自定义字符串错误码继续可用。

```ts
import { createDocxKitError, serializeDocxKitError } from 'docx-kit'

const error = createDocxKitError('PLUGIN_NOT_REGISTERED', { pluginName: 'qrcode' })
console.error(serializeDocxKitError(error))
```

## 常见错误场景

### 导出失败

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

### 插件未注册

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

### 未知类名

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

### 无效图片数据

```ts
try {
  doc.image({ data: '' })  // empty data
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.IMAGE_INVALID_DATA) {
    console.error('Image data is empty.')
  }
}
```

### 插件渲染失败

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

## 最佳实践

### 类型安全的错误匹配

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

### 异步错误处理

所有导出方法（`save()`、`toBlob()`、`toBuffer()`、`toBase64()`）都是异步的，请始终使用 `await`：

```ts
// ✅ Correct
await doc.save('output.docx')

// ❌ Wrong — save() returns a Promise, not void
doc.save('output.docx')
```

### 封装批量生成

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
        ? `[${err.diagnosticCode ?? err.code}] ${err.message}`
        : String(err)
      results.push({ item, error: message })
    }
  }

  return results
}
```
