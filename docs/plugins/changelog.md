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
