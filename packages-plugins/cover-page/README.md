# @docxkit/plugin-cover-page

Cover page plugin for docx-kit.

## Usage

```ts
import { coverPagePlugin } from '@docxkit/plugin-cover-page'

builder.use(coverPagePlugin())
builder.coverPage({ title: 'Annual Report', subtitle: '2026' })
```
