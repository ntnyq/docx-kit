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

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `BadgeConfig` and `BadgeStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `text` | `TextStyleRule` | Run typography |
| `paragraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |

```ts
badgePlugin({
  styles: { text: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
