# Figure

Combines an inline image, an optional caption, and an optional source credit. It uses the core image compiler for format detection, dimensions, alternative text, and image resolution.

## Import

```ts
import { figurePlugin, type FigureOptions } from 'docx-kit'
```

The standalone package exports the same factory and type from `@docxkit/plugin-figure`.

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `data` | `string \| Uint8Array \| ArrayBuffer \| Blob` | Required | Image bytes, Blob, data URL, raw base64, or a resolved path |
| `caption` | `string` | — | Literal caption below the image |
| `source` | `string` | — | Literal credit below the caption; no prefix is added |
| `alt` | `string` | — | Accessible image description |
| `width` | `UnitValue` | Intrinsic width | Bare numbers are pixels; CSS length units are supported |
| `height` | `UnitValue` | Intrinsic height | One supplied dimension preserves the image's aspect ratio |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'` | Detected | Explicit image format hint |
| `alignment` | `'center' \| 'left' \| 'right'` | `'center'` | Alignment of the image, caption, and credit |

When both dimensions are supplied they are used as given. If intrinsic dimensions cannot be read, the core image compiler uses its 300 × 180 pixel fallback and aspect ratio. Missing captions and credits produce no extra paragraphs.

## Example

```ts
import { createDocx, figurePlugin } from 'docx-kit/node'

await createDocx()
  .use(figurePlugin())
  .plugin('figure', {
    data: './architecture.png',
    width: '12cm',
    alt: 'Services and their dependencies',
    caption: 'Figure 1. Service architecture',
    source: 'Source: Engineering team',
    alignment: 'center',
  })
  .save('architecture.docx')
```

For browsers, import from `docx-kit`, supply bytes, a Blob, or a data URL, and export with `.toBlob()`. File paths require `docx-kit/node` or an explicit `config.resolveImage` adapter. URLs also require an explicit resolver; the plugin does not fetch them automatically. See [Images](/guide/images).

Image and caption paragraphs request keeping with the next paragraph when a caption or credit follows. Invocation styles can override this and the default alignment. Credits use the theme's `muted` color.

Caption text is rendered as supplied. Automatic numbering, Word sequence fields, and cross-references are outside this plugin's current scope.
