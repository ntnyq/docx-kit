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
import { createDocx, calloutPlugin } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin)
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
  .use(calloutPlugin)
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
  .use(calloutPlugin)
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
  .use(calloutPlugin)
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
  .use(calloutPlugin)
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
