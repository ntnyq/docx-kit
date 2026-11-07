# Changelog

Generate a release-note section with a title and a 4-column version table.

## Import

```ts
import { changelogPlugin, type ChangelogOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `entries` | `ChangelogEntry[]` | required | List of version rows |
| `title` | `string` | `'Changelog'` | Heading text shown above the table |

`ChangelogEntry` fields:

| Field | Type | Description |
| --- | --- | --- |
| `version` | `string` | Version label |
| `date` | `string` | Release date |
| `type` | `'added' \| 'changed' \| 'fixed' \| 'removed'` | Entry category with its own highlight color |
| `changes` | `string` | Human-readable change summary |

## Example

```ts
import { changelogPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(changelogPlugin())
  .plugin('changelog', {
    title: 'Release Notes',
    entries: [
      { version: '1.4.0', date: '2026-06-14', type: 'added', changes: 'Added zh-TW docs locale.' },
      { version: '1.4.1', date: '2026-06-15', type: 'fixed', changes: 'Corrected plugin sidebar links.' },
    ],
  })
```

## Notes

- Empty `entries` still render the title with a centered `(no entries)` placeholder.
- The plugin is a good fit for release notes, migration logs, or compliance change histories.

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `ChangelogConfig` and `ChangelogStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `bodyText`, `emptyText`, `headerText`, `title` | `TextStyleRule` | Run typography |
| `bodyParagraph`, `emptyParagraph`, `headerParagraph`, `titleParagraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |
| `bodyCell`, `headerCell` | `Omit<CellStyleRule, 'height'>` | Cell shading, borders, width and padding |

```ts
changelogPlugin({
  styles: { headerText: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
