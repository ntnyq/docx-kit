# 数据表格（Data Table）

将对象数组渲染为带样式表格，根据第一个对象的键自动推断列。

## 导入

```ts
import { dataTablePlugin, type DataTableOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `data` | `Record<string, unknown>[]` | _必填_ | 要渲染的数据，每个对象对应一行 |
| `align` | `Record<string, 'left' \| 'center' \| 'right'>` | 自动 | 各列对齐方式，默认根据值类型检测 |
| `bordered` | `boolean` | `true` | 渲染可见表格边框 |
| `format` | `Record<string, 'currency' \| 'date' \| 'number' \| 'percent'>` | — | 各列的值格式化方式 |
| `labels` | `Record<string, string>` | — | 可读列标签，如 `{ salary: '薪资' }` |
| `striped` | `boolean` | `false` | 交替行背景底纹 |
| `styles` | `DataTableStyles` | — | 本次调用的部位样式 |

## 示例

### 基础数据表格

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

### 设置标签与格式

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

### 自定义对齐并隐藏边框

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

### 空数据处理

`data` 为空数组时，插件渲染居中占位文字：

```ts
const doc = createDocx()
  .use(dataTablePlugin())
  .h1('Search Results')
  .plugin('dataTable', { data: [] })
  .p('No results found matching your criteria.')
  .save('empty.docx')
```

### 带斑马纹的大数据集

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

## 样式定制

给插件工厂传入 `{ styles }`，可设置当前插件实例的默认样式；通过 `options.styles` 覆盖单次调用。插件包和 `docx-kit` 的 Node/browser 入口均导出 `CalloutConfig` / `CalloutStyles` 与 `DataTableConfig` / `DataTableStyles`。

优先级从低到高：**内置默认值（含主题取值）→ 初始化 `styles` → 调用通用 `style`（`.plugin()` 第三个参数）→ 调用部位 `styles`**。缺省或 `undefined` 字段沿用默认值；嵌套普通对象按字段合并，数组整体替换，`0` 和 `false` 均生效。每次调用独立解析，不修改配置对象。单位和 `$colors.primary`、`$fontSize.lg`、`$spacing.md` 等主题 token 复用现有样式编译器。

| 部位 | 样式类型 | 作用范围 |
| --- | --- | --- |
| `headerCell`、`bodyCell` | `Omit<CellStyleRule, 'height'>` | 表头及数据单元格的背景、边框、宽度与内边距 |
| `stripedCell` | `Omit<CellStyleRule, 'height'>` | 斑马纹数据单元格的附加样式 |
| `headerParagraph`、`bodyParagraph`、`emptyParagraph` | `ParagraphStyleRule` | 表头、正文、空数据提示的段落样式 |
| `headerText`、`bodyText`、`emptyText` | `TextStyleRule` | 表头、正文、空数据提示的文字样式 |

`bodyCell` 应用于所有数据行；启用 `striped: true` 时，第 2、4 等数据行额外应用 `stripedCell`。同一优先级内先应用 `bodyCell`，再应用 `stripedCell`；调用级 `bodyCell` 仍高于初始化 `stripedCell`。显式样式边框可覆盖 `bordered` 选定的默认值，包括 `bordered: false`。段落部位样式可覆盖自动推导或按列设置的对齐。空数据仅使用 `emptyParagraph`、`emptyText`。

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

### JSON DSL、AI 与 MCP

JSON DSL 中在调用的 `options.styles` 内填写与 Builder 相同的可序列化部位样式。AI 工具定义和 schema 指南已包含这些选项；MCP 通过 `docx-kit://schema` 和 `get_plugin_help().optionsSchema` 暴露对应结构。`validate_schema` 会拒绝未知部位、放错部位的属性及错误的样式值。其他内置插件也提供对应的部位样式契约。

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
