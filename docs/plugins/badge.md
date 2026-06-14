# Badge

Render short status labels such as `DRAFT`, `APPROVED`, or `INTERNAL`.

## Import

```ts
import { badgePlugin, type BadgeOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `text` | `string` | required | Badge label text |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | `'neutral'` | Preset tone name or custom text color |
| `backgroundColor` | `string` | preset-based | Custom badge background color |

## Example

```ts
import { badgePlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(badgePlugin())
  .plugin('badge', {
    text: 'DRAFT',
    color: 'warning',
  })
```

## Notes

- Preset colors ship with matching foreground/background pairs.
- Custom `color` values are treated as literal text colors when they are not one of the preset names.
