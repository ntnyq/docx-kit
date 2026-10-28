# @docxkit/plugin-checklist

Renders acceptance criteria, inspection records, and release tasks with native Word checkbox controls. Items remain in input order.

## Usage

```ts
import { createDocx } from '@docxkit/core'
import { checklistPlugin } from '@docxkit/plugin-checklist'

const doc = createDocx()
  .use(checklistPlugin())
  .plugin('checklist', {
    title: 'Release checklist',
    items: [
      {
        text: 'Review changes',
        checked: true,
        owner: 'Alice',
        group: 'Before release',
      },
      {
        text: 'Acceptance tests',
        owner: 'Bob',
        note: 'Attach test report',
        group: 'Before release',
      },
      { text: 'Monitor errors', group: 'After release' },
    ],
  })

await doc.toBlob()
```

## Options

| Option  | Type              | Default  | Description                                        |
| ------- | ----------------- | -------- | -------------------------------------------------- |
| `items` | `ChecklistItem[]` | Required | Ordered tasks; an empty array renders `(no items)` |
| `title` | `string`          | —        | Optional checklist heading                         |

### `ChecklistItem`

| Field     | Type      | Default  | Description                              |
| --------- | --------- | -------- | ---------------------------------------- |
| `text`    | `string`  | Required | Task label and checkbox alias            |
| `checked` | `boolean` | `false`  | Initial checked state                    |
| `group`   | `string`  | —        | Heading emitted when the group changes   |
| `owner`   | `string`  | —        | Responsible person, after the task label |
| `note`    | `string`  | —        | Indented note below the task             |

Adjacent items with the same group share one heading. A group that reappears later gets a new heading; items are never sorted or regrouped. Ungrouped items reset the current group.

See the [full reference](../../docs/plugins/checklist.md) for platform notes and styling.
