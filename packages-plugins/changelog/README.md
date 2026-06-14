# @docxkit/plugin-changelog

Version changelog table plugin for docx-kit.

## Usage

```ts
import { changelogPlugin } from '@docxkit/plugin-changelog'

builder.use(changelogPlugin())
builder.plugin('changelog', {
  entries: [
    { version: '1.0.0', date: '2026-06', changes: 'Initial release', type: 'added' },
  ],
})
```

## Notes

- Renders a title paragraph followed by a 4-column changelog table
- Supported entry types: `added`, `changed`, `fixed`, `removed`
