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

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `DividerConfig` and `DividerStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `paragraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |

```ts
dividerPlugin({
  styles: { paragraph: { borderBottom: { color: '#2563EB', width: 2 }, marginTop: 0 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
