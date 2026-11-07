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

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `MetricCardConfig` 和 `MetricCardStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `change`, `emptyText`, `label`, `unit`, `value` | `TextStyleRule` | 文字排版 |
| `changeParagraph`, `emptyParagraph`, `labelParagraph`, `paddingParagraph`, `valueParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |
| `cell` | `Omit<CellStyleRule, 'height'>` | 单元格背景、边框、宽度和内边距 |

```ts
metricCardPlugin({
  styles: { value: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
