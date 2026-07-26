# Watermark

Renders a large, semi-transparent text watermark as a styled paragraph. Best placed in a section header or footer.

## Import

```ts
import { watermarkPlugin, type WatermarkOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `text` | `string` | _(required)_ | Watermark text |
| `alignment` | `AlignmentType` | `CENTER` | Horizontal alignment |
| `color` | `string` | `'BFBFBF'` | Text color in hex RRGGBB |
| `fontSize` | `number` | `48` | Font size in half-points (48 = 24pt) |

## Examples

### Basic Watermark

```ts
import { createDocx, watermarkPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Confidential Document')
  .p('This document contains proprietary information.')
  .plugin('watermark', { text: 'CONFIDENTIAL' })
  .save('confidential.docx')
```

### Red DRAFT Watermark

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Work in Progress')
  .p('This document is still under review.')
  .plugin('watermark', {
    text: 'DRAFT',
    color: 'FF0000',
    fontSize: 72, // 36pt
  })
  .save('draft.docx')
```

### Left-Aligned Watermark

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Internal Memo')
  .plugin('watermark', {
    text: 'INTERNAL USE ONLY',
    alignment: AlignmentType.LEFT,
    color: '999999',
  })
  .save('internal.docx')
```

### In Section Header/Footer

For a true document-wide watermark, place it in a header or footer:

```ts
const doc = createDocx()
  .use(watermarkPlugin())

// Render watermark in a footer that applies to all pages
const wm = doc.plugin('watermark', {
  text: 'SAMPLE',
  color: 'D0D0D0',
  fontSize: 60,
})

// Note: The exact header/footer integration depends on
// how sections are configured in your document.
```

### Multiple Watermarks

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Review Document')
  .plugin('watermark', { text: 'DRAFT', color: 'FF6600' })
  .p('Please review the following content carefully.')
  .pageBreak()
  .plugin('watermark', { text: 'CONFIDENTIAL', color: 'CC0000' })
  .p('This section contains confidential financial data.')
  .save('multi-watermark.docx')
```
