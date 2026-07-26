# @docxkit/plugin-callout

Callout/admonition plugin for docx-kit.

## Usage

```ts
import { calloutPlugin } from '@docxkit/plugin-callout'

builder.use(calloutPlugin())
builder.callout({ type: 'info', text: 'Note: this is important.' })
```
