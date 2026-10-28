# @docxkit/plugin-figure

Combines an inline image, an optional caption, and an optional source credit. It uses the core image compiler for format detection, dimensions, alternative text, and image resolution.

## Usage

```ts
import { createDocx } from '@docxkit/core'
import { figurePlugin } from '@docxkit/plugin-figure'

const image =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
await createDocx()
  .use(figurePlugin())
  .plugin('figure', {
    data: image,
    width: 120,
    caption: 'Figure 1. Sample image',
    source: 'Source: Example',
  })
  .toBlob()
```

## Options

| Option      | Type                                          | Default          | Description                                                 |
| ----------- | --------------------------------------------- | ---------------- | ----------------------------------------------------------- |
| `data`      | `string \| Uint8Array \| ArrayBuffer \| Blob` | Required         | Image bytes, Blob, data URL, raw base64, or a resolved path |
| `caption`   | `string`                                      | —                | Literal caption below the image                             |
| `source`    | `string`                                      | —                | Literal credit below the caption; no prefix is added        |
| `alt`       | `string`                                      | —                | Accessible image description                                |
| `width`     | `UnitValue`                                   | Intrinsic width  | Bare numbers are pixels; CSS length units are supported     |
| `height`    | `UnitValue`                                   | Intrinsic height | One supplied dimension preserves the image's aspect ratio   |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'`  | Detected         | Explicit image format hint                                  |
| `alignment` | `'center' \| 'left' \| 'right'`               | `'center'`       | Alignment of the image, caption, and credit                 |

When both dimensions are supplied they are used as given. If intrinsic dimensions cannot be read, the core image compiler uses its 300 × 180 pixel fallback and aspect ratio. Missing captions and credits produce no extra paragraphs.

See the [full reference](../../docs/plugins/figure.md) for platform notes and styling.
