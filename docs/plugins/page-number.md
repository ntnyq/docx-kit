# Page Number

Generates a paragraph containing a page number field for placement inside section headers or footers.

## Import

```ts
import { pageNumberPlugin, type PageNumberOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `alignment` | `AlignmentType` | `CENTER` | Horizontal alignment of the page number |
| `fontSize` | `number` | `20` | Font size in half-points (20 = 10pt) |
| `showTotal` | `boolean` | `false` | Show "Page X of Y" instead of just "X" |

## Examples

### Simple Page Number

```ts
import { createDocx, pageNumberPlugin } from 'docx-kit'

const doc = createDocx()
  .use(pageNumberPlugin)
  .plugin('pageNumber', {})
  .save('page-number.docx')
```

### "Page X of Y" Format

```ts
const doc = createDocx()
  .use(pageNumberPlugin)
  .h1('Report')
  .p('This report contains multiple pages.')
  .plugin('pageNumber', {
    showTotal: true,
  })
  .save('page-x-of-y.docx')
```

### Left-Aligned Page Number

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(pageNumberPlugin)
  .h1('Document')
  .plugin('pageNumber', {
    alignment: AlignmentType.LEFT,
    fontSize: 18, // 9pt
  })
  .save('left-pn.docx')
```

### Custom Font Size

```ts
const doc = createDocx()
  .use(pageNumberPlugin)
  .h1('Large Print Document')
  .plugin('pageNumber', {
    fontSize: 28, // 14pt
    showTotal: true,
  })
  .save('large-pn.docx')
```

### Integration with Headers/Footers

The page number plugin is best used within section headers/footers. Here's a conceptual example:

```ts
// Note: The plugin output (a Paragraph) can be used in section
// header/footer configuration. The exact API depends on how
// sections with headers/footers are set up.
```
