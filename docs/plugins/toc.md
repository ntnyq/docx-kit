# Table of Contents

Insert a Word table-of-contents field driven by heading levels.

## Import

```ts
import { tocPlugin, type TocOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `title` | `string` | `'Contents'` | Visible section title |
| `maxLevel` | `number` | `3` | Maximum heading level included, clamped to `1..9` |

## Example

```ts
import { createDocx, tocPlugin } from 'docx-kit'

const doc = createDocx()
  .use(tocPlugin())
  .plugin('toc', { title: 'Document Contents', maxLevel: 4 })
```

## Notes

- The plugin emits a real DOCX TOC field, not a static text snapshot.
- Microsoft Word may ask the user to update fields after opening the file.
