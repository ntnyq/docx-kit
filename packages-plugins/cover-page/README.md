# @docxkit/plugin-cover-page

Cover page plugin for docx-kit.

## Usage

```ts
import { coverPagePlugin } from '@docxkit/plugin-cover-page'

builder.use(coverPagePlugin())
builder.plugin('coverPage', { title: 'Annual Report', subtitle: '2026' })
```

The cover inherits the document font and uses Word's native Title style. A decorative rule is off by default; enable it with `showRule: true`. Call `.pageBreak()` after the plugin when the following content should start on a new page.
