# Example: Rich Content

Showcase inline rich content in paragraphs: mixed text runs, inline images, and the `span()` / `inlineImg()` DSL helpers.

## Full Code

```ts
import { createDocx, defineStyles, span, inlineImg } from 'docx-kit/node'

// 1. Styles
const styles = defineStyles({
  body: {
    fontSize: 12,
    lineHeight: 1.6,
  },
  caption: {
    fontSize: 10,
    color: '#888',
    textAlign: 'center',
    fontStyle: 'italic',
  },
})

// 2. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
})

// A small inline icon as a base64 data URL (1x1 px transparent gif placeholder)
const iconDataUrl =
  'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'

doc
  .h1('Rich Inline Content')
  .p('This example demonstrates mixed text runs and inline images within a paragraph.')

  // ─── Mixed text runs via children array ─────────────────────────────
  .p([
    span('This sentence has '),
    span('bold red text', { bold: true, color: '#e11d48' }),
    span(', '),
    span('italic blue text', { italic: true, color: '#2563eb' }),
    span(', and '),
    span('highlighted text', { highlight: 'yellow' }),
    span(' — all in the same paragraph.'),
  ])

  // ─── Inline image inside paragraph text ────────────────────────────
  .p([
    span('Here is an inline icon: '),
    inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
    span(' and some text after it.'),
  ])

  // ─── Styled link-like text ─────────────────────────────────────────
  .p([
    span('Visit our site: '),
    span('https://example.com', {
      color: '#2563eb',
      underline: true,
      bold: true,
    }),
    span(' for more info.'),
  ])

  // ─── Nested emphasis ──────────────────────────────────────────────
  .p([
    span('Note: '),
    span('This is important!', { bold: true, color: '#dc2626' }),
    span(' Please review before proceeding.'),
  ])

  .save('rich-content-demo.docx')
```

## What This Demonstrates

| Feature | Used In |
|---|---|
| `span(text, style?)` helper | Creating inline text runs with per-span styles |
| `inlineImg(opts)` helper | Inline images embedded in paragraph flow |
| `ParagraphNode.children` as array | Mixed content in a single paragraph |
| Per-span `bold` / `italic` / `color` | Inline style overrides within paragraph |
| Per-span `highlight` | Text highlighting within paragraph |
| Per-span `underline` | Underlined text within paragraph |

## `span()` Helper

```ts
import { span } from 'docx-kit'

// Plain text span
span('Hello ')

// Styled text span
span('world', { bold: true, color: '#f00' })

// Used inside p() children array
doc.p([
  span('Hello '),
  span('world', { bold: true }),
  span('!'),
])
```

## `inlineImg()` Helper

```ts
import { inlineImg } from 'docx-kit'

// Inline image from URL, base64, Buffer, or Uint8Array
inlineImg({
  data: 'https://example.com/icon.png',
  width: 16,
  height: 16,
})

// Used inside p() children array
doc.p([
  span('See icon: '),
  inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
  span(' done!'),
])
```

## Paragraph Children API

```ts
// String (classic API — all text in one run)
doc.p('Hello world', { className: 'body' })

// TextNode array (fine-grained control)
doc.p([
  { text: 'Hello ', type: 'text' },
  { text: 'world', type: 'text', style: { bold: true } },
])

// Mixed: text + inline image
doc.p([
  { text: 'Icon: ', type: 'text' },
  { type: 'image', data: imgData, width: 16, height: 16 },
  { text: ' done!', type: 'text' },
])

// With span() / inlineImg() helpers (recommended — cleaner syntax)
doc.p([
  span('Icon: '),
  inlineImg({ data: imgData, width: 16, height: 16 }),
  span(' done!'),
])
```
