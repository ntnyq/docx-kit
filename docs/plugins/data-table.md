# Data Table

Renders an array of objects as a styled table. Columns are auto-inferred from the first object's keys.

## Import

```ts
import { dataTablePlugin, type DataTableOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `data` | `Record<string, unknown>[]` | _(required)_ | The data to render — each object is one row |
| `align` | `Record<string, 'left' \| 'center' \| 'right'>` | auto | Per-column alignment. Auto-detected from value types |
| `bordered` | `boolean` | `true` | Render visible table borders |
| `format` | `Record<string, 'currency' \| 'date' \| 'number' \| 'percent'>` | — | Per-column value formatter |
| `labels` | `Record<string, string>` | — | Human-readable column labels (e.g. `{ salary: '薪资' }`) |
| `striped` | `boolean` | `false` | Alternate row background shading |

## Examples

### Basic Data Table

```ts
import { createDocx, dataTablePlugin } from 'docx-kit'

const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Employee Directory')
  .plugin('dataTable', {
    data: [
      { name: 'Alice Chen', department: 'Engineering', role: 'Senior Developer' },
      { name: 'Bob Wang', department: 'Design', role: 'UX Designer' },
      { name: 'Carol Li', department: 'Engineering', role: 'Tech Lead' },
    ],
  })
  .save('employees.docx')
```

### With Labels and Formatting

```ts
const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Financial Summary')
  .plugin('dataTable', {
    data: [
      { name: 'Product A', revenue: 850000, growth: 0.15, launchDate: new Date('2025-03-15') },
      { name: 'Product B', revenue: 620000, growth: 0.08, launchDate: new Date('2025-06-01') },
      { name: 'Product C', revenue: 1200000, growth: 0.23, launchDate: new Date('2025-01-10') },
    ],
    labels: {
      name: '产品名称',
      revenue: '收入',
      growth: '增长率',
      launchDate: '上市日期',
    },
    format: {
      revenue: 'currency',
      growth: 'percent',
      launchDate: 'date',
    },
    striped: true,
  })
  .save('financial.docx')
```

### Custom Alignment and No Borders

```ts
const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Project Metrics')
  .plugin('dataTable', {
    data: [
      { metric: 'Page Load Time', value: '1.2s', target: '< 2s' },
      { metric: 'API Response', value: '45ms', target: '< 100ms' },
      { metric: 'Error Rate', value: '0.05%', target: '< 1%' },
      { metric: 'Uptime', value: '99.99%', target: '> 99.9%' },
    ],
    align: { value: 'center', target: 'center' },
    bordered: false,
    striped: true,
  })
  .save('metrics.docx')
```

### Empty Data Handling

When `data` is an empty array, the plugin renders a centered placeholder text:

```ts
const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Search Results')
  .plugin('dataTable', { data: [] })
  .p('No results found matching your criteria.')
  .save('empty.docx')
```

### Large Dataset with Striped Rows

```ts
const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const monthlyData = monthNames.map((month, i) => ({
  month,
  revenue: Math.round(800000 + Math.random() * 400000),
  expenses: Math.round(500000 + Math.random() * 300000),
  profit: 0, // computed below
}))

// Compute profit
monthlyData.forEach(d => { d.profit = d.revenue - d.expenses })

const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Monthly Financial Report')
  .plugin('dataTable', {
    data: monthlyData,
    labels: { month: 'Month', revenue: 'Revenue', expenses: 'Expenses', profit: 'Profit' },
    format: { revenue: 'currency', expenses: 'currency', profit: 'currency' },
    striped: true,
    align: { month: 'left', revenue: 'right', expenses: 'right', profit: 'right' },
  })
  .save('monthly-report.docx')
```
