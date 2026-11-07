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

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `TocConfig` and `TocStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `title` | `TextStyleRule` | Run typography |
| `titleParagraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |

Only the explicit title is styled here. Word generates TOC entries using its TOC document styles when updating the field.

```ts
tocPlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
