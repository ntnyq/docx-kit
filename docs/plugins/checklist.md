# Checklist

Renders acceptance criteria, inspection records, and release tasks with native Word checkbox controls. Items remain in input order.

## Import

```ts
import { checklistPlugin, type ChecklistItem, type ChecklistOptions } from 'docx-kit'
```

The standalone package exports the same factory and types from `@docxkit/plugin-checklist`.

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `items` | `ChecklistItem[]` | Required | Ordered tasks; an empty array renders `(no items)` |
| `title` | `string` | — | Optional checklist heading |

### `ChecklistItem`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `text` | `string` | Required | Task label and checkbox alias |
| `checked` | `boolean` | `false` | Initial checked state |
| `group` | `string` | — | Heading emitted when the group changes |
| `owner` | `string` | — | Responsible person, after the task label |
| `note` | `string` | — | Indented note below the task |

Adjacent items with the same group share one heading. A group that reappears later gets a new heading; items are never sorted or regrouped. Ungrouped items reset the current group.

## Example

```ts
import { checklistPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(checklistPlugin())
  .plugin('checklist', {
    title: 'Release checklist',
    items: [
      { text: 'Review changes', checked: true, owner: 'Alice', group: 'Before release' },
      { text: 'Run acceptance tests', owner: 'Bob', note: 'Attach the test report', group: 'Before release' },
      { text: 'Monitor errors', group: 'After release' },
    ],
  })

await doc.toBlob()
```

The controls can be toggled in Word readers that support checkbox content controls. Browser previews may display only their initial state. Notes and owners use the theme's `muted` color; invocation styles override default text and paragraph styling.
