# @docxkit/plugin-metric-card

Renders report indicators in an equal-width table grid. Each card shows a label, a prominent value, and optional unit and comparison text.

## Usage

```ts
import { createDocx } from '@docxkit/core'
import { metricCardPlugin } from '@docxkit/plugin-metric-card'

const doc = createDocx()
  .use(metricCardPlugin())
  .plugin('metricCard', {
    columns: 3,
    items: [
      {
        label: 'Revenue',
        value: '128,000',
        unit: 'USD',
        trend: 'up',
        change: '+12%',
      },
      { label: 'Active users', value: 4200, trend: 'down', change: '-3%' },
      {
        label: 'Uptime',
        value: 99.9,
        unit: '%',
        trend: 'neutral',
        change: 'Unchanged',
      },
      { label: 'Incidents', value: 0 },
    ],
  })

await doc.toBlob()
```

## Options

| Option    | Type           | Default  | Description                                                     |
| --------- | -------------- | -------- | --------------------------------------------------------------- |
| `items`   | `MetricItem[]` | Required | Indicators in display order; empty arrays render `(no metrics)` |
| `columns` | `number`       | `3`      | Maximum cards per row; must be a positive safe integer          |

### `MetricItem`

| Field    | Type                          | Description                                                   |
| -------- | ----------------------------- | ------------------------------------------------------------- |
| `label`  | `string`                      | Indicator name                                                |
| `value`  | `number \| string`            | Display value, including zero; no automatic locale formatting |
| `unit`   | `string`                      | Unit appended after a space                                   |
| `change` | `string`                      | Literal comparison text such as `+12% vs last month`          |
| `trend`  | `'up' \| 'down' \| 'neutral'` | Optional arrow and color: ↑ success, ↓ danger, → muted        |

`change` and `trend` work independently. No trend is inferred from the value or change text. When a decrease is desirable, omit `trend` and provide your own comparison text or override the invocation color.

See the [full reference](../../docs/plugins/metric-card.md) for platform notes and styling.
