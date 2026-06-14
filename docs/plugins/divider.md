# Divider

Insert a horizontal rule between sections or visually separate content blocks.

## Import

```ts
import { dividerPlugin, type DividerOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | `'solid'` | Border style for the separator |
| `color` | `string` | `'D9D9D9'` | Divider color in hex |
| `spacingBefore` | `number` | `200` | Paragraph spacing before the rule |
| `spacingAfter` | `number` | `200` | Paragraph spacing after the rule |

## Example

```ts
import { createDocx, dividerPlugin } from 'docx-kit'

const doc = createDocx()
  .use(dividerPlugin())
  .h1('Quarterly Review')
  .plugin('divider', { style: 'double', color: '4472C4' })
  .p('Executive summary starts here.')
```
