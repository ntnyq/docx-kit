# 指标卡（Metric Card）

以等宽表格网格渲染报告指标。每张卡包含标签、醒目的数值，以及可选单位和对比文字。

## 导入

```ts
import { metricCardPlugin, type MetricCardOptions, type MetricItem } from 'docx-kit'
```

独立包 `@docxkit/plugin-metric-card` 导出相同工厂函数与类型，调用名称为 `metricCard`。

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `items` | `MetricItem[]` | 必填 | 按显示顺序排列的指标；空数组渲染 `(no metrics)` |
| `columns` | `number` | `3` | 每行最多卡片数，必须为正安全整数 |

### `MetricItem`

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `label` | `string` | 指标名称 |
| `value` | `number \| string` | 显示值，包括零；不自动按地区格式化 |
| `unit` | `string` | 在空格后追加的单位 |
| `change` | `string` | 原样输出的对比文字，如 `+12% vs last month` |
| `trend` | `'up' \| 'down' \| 'neutral'` | 可选箭头与颜色：↑ 成功色，↓ 危险色，→ 弱化色 |

`change` 和 `trend` 独立工作，不会根据数值或变化文字推断趋势。下降代表好转时，可以省略 `trend` 并提供自己的对比文字，或覆盖调用颜色。

## 示例

```ts
import { createDocx, metricCardPlugin } from 'docx-kit'

const doc = createDocx()
  .use(metricCardPlugin())
  .plugin('metricCard', {
    columns: 3,
    items: [
      { label: 'Revenue', value: '128,000', unit: 'USD', trend: 'up', change: '+12% vs last month' },
      { label: 'Active users', value: 4200, trend: 'down', change: '-3%' },
      { label: 'Uptime', value: 99.9, unit: '%', trend: 'neutral', change: 'Unchanged' },
      { label: 'Incidents', value: 0 },
    ],
  })

await doc.toBlob()
```

每行不会跨页拆分。最后一行补齐空单元格；条目少于 `columns` 时，现有卡片平分全部宽度。无效列数会抛出代码为 `PLUGIN_RENDER_FAILED` 的 `DocxKitError`，空输入也会检查列数。

网格使用主题的 `surface`、`primary`、`muted`、`success` 和 `danger` 颜色，并提供内置回退值。调用样式可覆盖文字、段落、单元格背景与边框。
