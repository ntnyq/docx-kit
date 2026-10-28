# Checklist

Renders acceptance criteria, inspection records, and release tasks with GitHub-style graphic checkboxes by default. Items remain in input order.

## Import

```ts
import { checklistPlugin, type ChecklistItem, type ChecklistOptions } from 'docx-kit'
```

The standalone package exports the same factory and types from `@docxkit/plugin-checklist`.

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `items` | `ChecklistItem[]` | Required | Ordered tasks; an empty array renders `(no items)` |
| `checkboxStyle` | `'github' \| 'native'` | `'github'` | Static rounded graphics, or native Word controls that can be toggled |
| `title` | `string` | — | Optional checklist heading |

### `ChecklistItem`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `text` | `string` | Required | Task label and checkbox alias |
| `checked` | `boolean` | `false` | Initial checked state |
| `group` | `string` | — | Heading emitted when the group changes |
| `owner` | `string` | — | Responsible person, after the task label |
| `note` | `string` | — | Indented note below the task |

Adjacent items with the same group share one heading. A group that reappears later gets a new heading; items are never sorted or regrouped. Ungrouped items reset the current group. Group headings are bold, use the theme's `accent` color (falling back to `primary`), and scale to approximately 120% of the task font size. The optional checklist title uses `primary` at approximately 140%, preserving a clear hierarchy above the groups. Explicit invocation colors override these defaults.

## Complete example

[Open this example in the Playground](/playground?example=checklist) to edit the code, preview it, or download the DOCX.

This Ocean theme sample includes both groups, owners, notes, and checked and unchecked tasks. Body text is 12pt, group headings are 14.5pt, owners are 10pt, and notes are 9.5pt.

![Complete checklist with group headings, owners, and notes](/examples/checklist.png)

<<< @/.vitepress/constants/templates/checklist.ts#checklist-example

The default `'github'` style uses SVG graphics with high-resolution PNG fallbacks. Checked boxes use the theme's `accent` color (falling back to `primary`) with a white check. Unchecked outlines use `muted` (falling back to `border`), and their fill uses `background`. Both image formats follow the same theme colors. Graphic checkboxes scale with the task's effective font size and align with the text's visual center; wrapped text and notes share the same text inset. These are static images and do not depend on symbol fonts.

Set `checkboxStyle: 'native'` to use Word checkbox content controls that can be toggled in supported readers. Browser previews may display only their initial state. Notes and owners use the theme's `muted` color; invocation styles override default text and paragraph styling. Notes scale to approximately 80% of the task's effective font size, rounded to Word's half-point precision. Owner labels scale to approximately 85% and retain the theme's `muted` color even when task text has an explicit color override.
