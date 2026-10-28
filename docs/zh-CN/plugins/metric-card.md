# Metric Card

以等宽多列表格展示指标摘要，调用名为 `metricCard`。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `items` | `MetricItem[]` | 指标数组；空数组显示 `(no metrics)` |
| `columns` | `number` | 每行最多列数，默认 3，必须是正安全整数 |

`MetricItem` 包含必填的 `label`、`value: number | string`，以及可选的 `unit`、`change`、`trend: 'up' | 'down' | 'neutral'`。

数值按原样显示，保留 0，不做自动格式化；单位前添加空格。`change` 是比较文本，`trend` 显示 ↑（success）、↓（danger）、→（muted），两者可独立使用，不自动推断趋势。下降代表改善时，可以省略趋势并自行提供比较文本，或覆盖调用颜色。

末行补空单元格，指标少于列数时平分全宽，整行尽量不跨页。非法列数会抛出 `PLUGIN_RENDER_FAILED`，空数组也会检查。使用主题的 `surface`、`primary`、`muted`、`success`、`danger` 颜色，调用样式可以覆盖。

独立包：`@docxkit/plugin-metric-card`，导出 `metricCardPlugin`、`MetricCardOptions`、`MetricItem`。主包 `docx-kit` 和 `docx-kit/node` 同样导出。

## 用法

```ts
import { createDocx, metricCardPlugin } from 'docx-kit'

const doc = createDocx().use(metricCardPlugin()).plugin('metricCard', {
  columns: 3,
  items: [
    { label: '营收', value: '128,000', unit: '元', trend: 'up', change: '+12%' },
    { label: '活跃用户', value: 4200, trend: 'down', change: '-3%' },
    { label: '可用性', value: 99.9, unit: '%', trend: 'neutral', change: '持平' },
    { label: '事故', value: 0 },
  ],
})

await doc.toBlob()
```
