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
| `styles` | `DataTableStyles` | — | 本次呼叫的部位樣式 |


使用 `format: { column: 'date' }` 時，`YYYY-MM-DD` 字串在所有時區保留原始日曆日期。Date 物件與時間戳記字串仍依執行環境的本機時區顯示。無法解析的日期（包括無效 Date 物件）保留為可讀文字；純日期字串不會被悄悄正規化。

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

## 樣式自訂

傳入 `{ styles }` 給外掛工廠，可設定目前外掛實例的預設樣式；透過 `options.styles` 覆寫單次呼叫。外掛套件和 `docx-kit` 的 Node/browser 入口均匯出 `CalloutConfig` / `CalloutStyles` 與 `DataTableConfig` / `DataTableStyles`。

優先順序由低至高：**內建預設值（含佈景主題值）→ 初始化 `styles` → 呼叫通用 `style`（`.plugin()` 第三個參數）→ 呼叫部位 `styles`**。省略或 `undefined` 欄位沿用預設值；巢狀普通物件逐欄位合併，陣列整體取代，`0` 和 `false` 均生效。每次呼叫獨立解析，不修改設定物件。單位和 `$colors.primary`、`$fontSize.lg`、`$spacing.md` 等主題 token 沿用現有樣式編譯器。

| 部位 | 樣式型別 | 作用範圍 |
| --- | --- | --- |
| `headerCell`、`bodyCell` | `Omit<CellStyleRule, 'height'>` | 表頭及資料儲存格的背景、框線、寬度與內距 |
| `stripedCell` | `Omit<CellStyleRule, 'height'>` | 斑馬紋資料儲存格的附加樣式 |
| `headerParagraph`、`bodyParagraph`、`emptyParagraph` | `ParagraphStyleRule` | 表頭、內文、空資料提示的段落樣式 |
| `headerText`、`bodyText`、`emptyText` | `TextStyleRule` | 表頭、內文、空資料提示的文字樣式 |

`bodyCell` 套用至所有資料列；啟用 `striped: true` 時，第 2、4 等資料列額外套用 `stripedCell`。同一優先層級內先套用 `bodyCell`，再套用 `stripedCell`；呼叫級 `bodyCell` 仍優先於初始化 `stripedCell`。明確設定的樣式框線可覆寫 `bordered` 選定的預設值，包括 `bordered: false`。段落部位樣式可覆寫自動推導或逐欄設定的對齊。空資料僅使用 `emptyParagraph`、`emptyText`。

```ts
import { createDocx, dataTablePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(dataTablePlugin({
    styles: {
      headerCell: { backgroundColor: '#EFF6FF' },
      headerText: { color: '#1E40AF', fontSize: 12 },
      bodyCell: { margin: '5pt 8pt' },
      bodyText: { fontSize: 11 },
      stripedCell: { backgroundColor: '#F8FAFC' },
    },
  }))
  .plugin('dataTable', {
    data: [{ name: 'Alice', score: 95 }, { name: 'Bob', score: 88 }],
    striped: true,
    styles: {
      headerText: { bold: false },
      bodyCell: { marginLeft: 0 },
      stripedCell: { backgroundColor: '#E0F2FE' },
    },
  })
```

### JSON DSL、AI 與 MCP

JSON DSL 在呼叫的 `options.styles` 內填寫與 Builder 相同的可序列化部位樣式。AI 工具定義和 schema 指南已包含這些選項；MCP 透過 `docx-kit://schema` 和 `get_plugin_help().optionsSchema` 提供對應結構。`validate_schema` 會拒絕未知部位、放錯部位的屬性及錯誤的樣式值。其他內建外掛也提供對應的部位樣式契約。

```json
{
  "type": "plugin",
  "name": "dataTable",
  "options": {
    "data": [{ "score": 95 }, { "score": 88 }],
    "striped": true,
    "styles": { "stripedCell": { "backgroundColor": "#EFF6FF" } }
  }
}
```
