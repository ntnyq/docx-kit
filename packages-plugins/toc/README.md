# @docxkit/plugin-toc

Table of Contents plugin for docx-kit.

## Usage

```ts
import { tocPlugin } from '@docxkit/plugin-toc'

builder.use(tocPlugin())
builder.plugin('toc', { title: 'Contents', maxLevel: 3 })
```

## Notes

- Inserts a DOCX table-of-contents field
- `maxLevel` controls the heading range included in the TOC
- Word may require a field update after opening the generated file
