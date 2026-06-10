# Error Handling

docx-kit throws structured `DocxKitError` instances with machine-readable error codes.

## DocxKitError

```ts
class DocxKitError extends Error {
  readonly code: string | ErrorCode  // Machine-readable code
  readonly cause: unknown            // Underlying error (if any)
  name: 'DocxKitError'               // Always "DocxKitError"
}
```

### Catching Errors

```ts
import { DocxKitError, ERROR_CODES } from 'docx-kit'

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

## Error Codes

| Constant | Code | Description |
|---|---|---|
| `ERROR_CODES.EXPORT_FAILED` | `'EXPORT_FAILED'` | Document export (save/pack) failed |
| `ERROR_CODES.IMAGE_INVALID_DATA` | `'IMAGE_INVALID_DATA'` | Image data is empty, corrupt, or unsupported |
| `ERROR_CODES.IMAGE_UNKNOWN_TYPE` | `'IMAGE_UNKNOWN_TYPE'` | Could not determine image format |
| `ERROR_CODES.PLUGIN_NOT_REGISTERED` | `'PLUGIN_NOT_REGISTERED'` | Plugin node referenced unregistered plugin |
| `ERROR_CODES.PLUGIN_RENDER_FAILED` | `'PLUGIN_RENDER_FAILED'` | Plugin's `render()` threw an error |
| `ERROR_CODES.STYLE_INVALID_UNIT` | `'STYLE_INVALID_UNIT'` | Invalid unit string in a style value |
| `ERROR_CODES.STYLE_UNKNOWN_CLASS` | `'STYLE_UNKNOWN_CLASS'` | `className` referenced a non-existent stylesheet key |
| `ERROR_CODES.TABLE_INVALID_COLUMNS` | `'TABLE_INVALID_COLUMNS'` | Table created with no columns |
| `ERROR_CODES.UNKNOWN_NODE_TYPE` | `'UNKNOWN_NODE_TYPE'` | Node type has no registered compiler |

## Common Error Scenarios

### Export Failed

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

### Plugin Not Registered

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

### Unknown Class Name

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

### Invalid Image Data

```ts
try {
  doc.image({ data: '' })  // empty data
} catch (err) {
  if (err instanceof DocxKitError && err.code === ERROR_CODES.IMAGE_INVALID_DATA) {
    console.error('Image data is empty.')
  }
}
```

### Plugin Render Failed

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

## Best Practices

### Type-safe Error Matching

```ts
import { DocxKitError, ERROR_CODES } from 'docx-kit'

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

### Async Error Handling

All export methods (`save()`, `toBlob()`, `toBuffer()`, `toBase64()`) are async — always `await` them:

```ts
// ✅ Correct
await doc.save('output.docx')

// ❌ Wrong — save() returns a Promise, not void
doc.save('output.docx')
```

### Wrapping for Batch Generation

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
