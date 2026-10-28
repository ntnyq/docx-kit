# 資料表格（Data Table）

將物件陣列呈現為帶樣式表格，根據第一個物件的鍵自動推斷列。

## 匯入

```ts
import { dataTablePlugin, type DataTableOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `data` | `Record<string, unknown>[]` | _必填_ | 要呈現的資料，每個物件對應一行 |
| `align` | `Record<string, 'left' \| 'center' \| 'right'>` | 自動 | 各欄對齊方式，預設根據值型別檢測 |
| `bordered` | `boolean` | `true` | 呈現可見表格邊框 |
| `format` | `Record<string, 'currency' \| 'date' \| 'number' \| 'percent'>` | — | 各欄的值格式化方式 |
| `labels` | `Record<string, string>` | — | 可讀欄標籤，如 `{ salary: '薪资' }` |
| `striped` | `boolean` | `false` | 交替行背景底紋 |

## 範例

### 基本資料表格

```ts
import { createDocx, dataTablePlugin } from 'docx-kit/node'

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

### 設置標籤與格式

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

### 自訂對齊並隱藏邊框

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

### 空資料處理

`data` 為空陣列時，外掛呈現居中預留位置文字：

```ts
const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Search Results')
  .plugin('dataTable', { data: [] })
  .p('No results found matching your criteria.')
  .save('empty.docx')
```

### 帶斑馬紋的大資料集

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
