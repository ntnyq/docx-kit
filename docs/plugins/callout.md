# Callout

Colored info / warning / success / danger callout boxes with icons.

## Import

```ts
import { calloutPlugin, type CalloutOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `content` | `string` | _(required)_ | Body text of the callout |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | _(required)_ | Callout style — controls icon, color, and tone |
| `title` | `string` | — | Optional bold title line placed before the content |
| `styles` | `CalloutStyles` | — | Per-part styles for this invocation |

### Type Presets

| Type | Icon | Background | Border Color |
|------|------|------------|--------------|
| `info` | ℹ️ | `#D6E4F0` | `#4472C4` |
| `warning` | ⚠️ | `#FFF2CC` | `#FFC000` |
| `success` | ✅ | `#E2F0D9` | `#70AD47` |
| `danger` | ⚠️ | `#FCE4D6` | `#FF0000` |

## Examples

### Information Callout

```ts
import { createDocx, calloutPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(calloutPlugin())
  .h1('System Maintenance')
  .plugin('callout', {
    type: 'info',
    content: 'The system will be upgraded tonight at 22:00. No downtime is expected.',
  })
  .save('callout.docx')
```

### Warning with Title

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Important Notice')
  .plugin('callout', {
    type: 'warning',
    title: '注意',
    content: '此操作不可撤销，请确认后再提交。',
  })
  .save('warning.docx')
```

### Success Confirmation

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Deployment Status')
  .plugin('callout', {
    type: 'success',
    title: 'Deployment Successful',
    content: 'Version 2.4.1 has been deployed to production. All health checks passed.',
  })
  .save('success.docx')
```

### Danger Alert

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Security Alert')
  .plugin('callout', {
    type: 'danger',
    title: 'Critical Vulnerability',
    content: 'A critical security vulnerability (CVE-2026-XXXX) has been detected. Immediate action required.',
  })
  .save('danger.docx')
```

### Multiple Callouts

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Weekly Summary')

  .plugin('callout', {
    type: 'success',
    title: 'Completed',
    content: 'Sprint goals met on time. 12 features shipped.',
  })

  .plugin('callout', {
    type: 'warning',
    title: 'At Risk',
    content: 'API rate limiting needs attention before next sprint.',
  })

  .plugin('callout', {
    type: 'info',
    title: 'Upcoming',
    content: 'Team retro scheduled for Friday at 3 PM.',
  })

  .save('multiple-callouts.docx')
```

## Style Customization

Pass `{ styles }` to the factory for defaults shared by that plugin instance. Set `options.styles` to override one invocation. `CalloutConfig` / `CalloutStyles` and `DataTableConfig` / `DataTableStyles` are exported from the plugin packages and `docx-kit` Node/browser entries.

Precedence, from lowest to highest: **built-in defaults (including theme values) → initialization `styles` → common invocation `style` (the third `.plugin()` argument) → invocation part `styles`**. Missing or `undefined` fields retain defaults; nested plain objects merge by field; arrays replace entirely; `0` and `false` apply. Styles are resolved per invocation without changing the configuration objects. Units and theme tokens such as `$colors.primary`, `$fontSize.lg` and `$spacing.md` use the existing style compiler.

| Part | Style type | Target |
| --- | --- | --- |
| `cell` | `Omit<CellStyleRule, 'height'>` | Container background, borders, width and padding |
| `paragraph` | `ParagraphStyleRule` | Shared paragraph alignment, spacing and flow |
| `title` | `TextStyleRule` | Title typography |
| `content` | `TextStyleRule` | Body typography |

```ts
import { calloutPlugin, createDocx } from 'docx-kit/node'

const doc = createDocx({
  theme: { colors: { primary: '#2563EB' }, spacing: { md: 8 } },
})
  .use(calloutPlugin({
    styles: {
      cell: { margin: '8pt 10pt', borderLeft: { width: 3 } },
      title: { fontSize: 16, color: '$colors.primary' },
      content: { fontSize: 11 },
    },
  }))
  .plugin('callout', {
    type: 'info',
    title: 'Note',
    content: 'Custom body text',
    styles: {
      cell: { marginBottom: 0 },
      title: { bold: false, color: '#DC2626' },
      paragraph: { lineHeight: 1.5 },
    },
  }, { fontFamily: 'Arial' })
```

### JSON DSL, AI and MCP

Only invocation `options.styles` belongs in JSON DSL. It uses the same serializable part objects as the builder. AI tool definitions and the schema guide describe these options; MCP exposes them in `docx-kit://schema` and `get_plugin_help().optionsSchema`. `validate_schema` rejects unknown parts, properties in the wrong part and invalid style values. Other built-in plugins also expose their own part-style contracts.

```json
{
  "type": "plugin",
  "name": "callout",
  "options": {
    "type": "info",
    "title": "Note",
    "content": "Body",
    "styles": { "title": { "color": "#2563EB" }, "cell": { "margin": 0 } }
  }
}
```
