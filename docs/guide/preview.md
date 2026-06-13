# Browser Preview

Render `.docx` files **directly in the browser** — no server, no file upload, no conversion step.

`@docxkit/renderer` brings Microsoft Word documents to the web using
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (DOM-based, pixel-faithful) with an optional Microsoft Office Online iframe fallback.

---

## Install

```sh
npm install @docxkit/renderer
```

> `docx-preview` (Apache-2.0, ~47 kB gzipped) and `jszip` are installed automatically.

---

## Quick Start

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

## Input Types

`preview.render(input)` auto-detects the input type:

| Type | Example | Notes |
|---|---|---|
| `Blob` | `await doc.toBlob()` | Most common — from `DocxBuilder.toBlob()` |
| `File` | `fileInput.files[0]!` | From `<input type="file">` |
| `ArrayBuffer` | `await blob.arrayBuffer()` | From `fetch()` response |
| `Uint8Array` | `new Uint8Array(buffer)` | For binary data sources |
| `string` (URL) | `'https://example.com/doc.docx'` | Fetched via `fetch()` |

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

## Renderer Modes

### DOM Renderer (default)

Uses `docx-preview` to render the DOCX as styled HTML elements inside the container.

- **Text is selectable and searchable** (DOM-based)
- Supports headers/footers, page breaks, tables, images
- Fidelity: high (pixel-faithful to Word layout)

```ts
const preview = createDocxPreview(container, {
  renderer: 'dom',  // default
  pageMode: 'paged',  // or 'continuous'
})
```

### Microsoft Office Online (iframe fallback)

Embeds Microsoft's free online viewer via an `<iframe>`.

- **Requires a publicly accessible URL** (Microsoft's servers must be able to fetch the file)
- 100% fidelity (Microsoft's own renderer)
- No programmatic control

```ts
const preview = createDocxPreview(container, {
  renderer: 'microsoft',
})

// Only URL strings are supported
await preview.render('https://example.com/public/doc.docx')
```

For non-URL inputs, the Microsoft renderer throws with a clear error:

```
[MICROSOFT_URL_REQUIRED]
Microsoft renderer requires a publicly accessible URL string.
Received Blob. Upload the DOCX to a public URL first,
or use renderer: 'dom'.
```

---

## Pagination

### Paged Mode (default)

Each page rendered as a discrete white box with shadow — looks like a real document.

```ts
const preview = createDocxPreview(container, {
  pageMode: 'paged',  // default
})
```

### Continuous Mode

Single scrolling document, no page breaks.

```ts
const preview = createDocxPreview(container, {
  pageMode: 'continuous',
})
```

---

## API Reference

### `createDocxPreview(container, options?)`

Creates a preview instance bound to a DOM container.

**Parameters:**
- `container: HTMLElement` — The target DOM element
- `options: DocxPreviewOptions` — Optional configuration

**Returns:** `DocxPreview`

---

### `DocxPreview` Instance

| Method | Description |
|---|---|
| `render(input)` | Render a DOCX input (Blob/File/ArrayBuffer/Uint8Array/URL) |
| `clear()` | Remove rendered content, keep instance alive |
| `destroy()` | Full cleanup (revokes URLs, clears DOM, marks destroyed) |

| Property | Type | Description |
|---|---|---|
| `container` | `readonly HTMLElement` | The bound container element |
| `currentInput` | `readonly DocxInput \| null` | Most recently rendered input |

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

## Options

All options from `docx-preview` can be passed through:

| Option | Type | Default | Description |
|---|---|---|---|
| `renderer` | `'dom' \| 'microsoft'` | `'dom'` | Rendering backend |
| `pageMode` | `'paged' \| 'continuous'` | `'paged'` | Page display mode |
| `className` | `string` | `'docx-kit-preview'` | CSS class on container |
| `inWrapper` | `boolean` | `true` | Add wrapper `<div>` around pages |
| `breakPages` | `boolean` | `true` | Page breaks (overrides `pageMode`) |
| `renderHeaders` | `boolean` | `false` | Render page headers |
| `renderFooters` | `boolean` | `false` | Render page footers |
| `renderFootnotes` | `boolean` | `false` | Render footnotes |
| `renderEndnotes` | `boolean` | `false` | Render endnotes |
| `ignoreFonts` | `boolean` | `false` | Don't load web fonts |
| `useBase64URL` | `boolean` | `false` | Use base64 images (vs ObjectURL) |
| `microsoftViewerUrl` | `string` | Office Online URL | Custom OOS server |

---

## Error Handling

All errors are `DocxKitError` instances with a structured `code`:

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

| Code | Trigger |
|---|---|
| `PREVIEW_INPUT_INVALID` | Input is `null`, `undefined`, or unrecognized type |
| `PREVIEW_FETCH_FAILED` | `fetch()` to URL failed (network or HTTP error) |
| `PREVIEW_RENDER_FAILED` | `docx-preview` `renderAsync()` threw |
| `MICROSOFT_URL_REQUIRED` | Microsoft renderer called with non-URL input |

---

## Limitations

- **Browser-only** — The renderer uses DOM APIs; it cannot run in Node.js
- **`docx-preview` limitations** — very large documents (200+ pages) may be slow; track changes and comments rendering is experimental
- **Microsoft mode** — requires a publicly accessible URL; documents are sent to Microsoft's servers
- **CJK fonts** — supported if the user's system has the fonts installed; `docx-preview` inherits browser font handling

---

## See Also

- [Examples: DOCX Preview](/examples/preview) — End-to-end example
- [Builder API](/guide/builder-api) — Generate documents to preview
- [Platforms](/guide/platforms) — Node.js `save()` vs browser `toBlob()`
