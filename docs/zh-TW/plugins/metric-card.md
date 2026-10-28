# Metric Card

以等寬多欄表格展示指標摘要，呼叫名稱為 `metricCard`。

## 選項

| 選項 | 型別 | 說明 |
| --- | --- | --- |
| `items` | `MetricItem[]` | 指標陣列；空陣列顯示 `(no metrics)` |
| `columns` | `number` | 每列最多欄數，預設 3，必須是正安全整數 |

`MetricItem` 包含必填的 `label`、`value: number | string`，以及可選的 `unit`、`change`、`trend: 'up' | 'down' | 'neutral'`。

數值按原樣顯示，保留 0，不做自動格式化；單位前加入空格。`change` 是比較文字，`trend` 顯示 ↑（success）、↓（danger）、→（muted），兩者可獨立使用，不自動推斷趨勢。下降代表改善時，可以省略趨勢並自行提供比較文字，或覆寫呼叫色彩。

最後一列補空儲存格，指標少於欄數時平分全寬，整列盡量不跨頁。無效欄數會拋出 `PLUGIN_RENDER_FAILED`，空陣列也會檢查。使用主題的 `surface`、`primary`、`muted`、`success`、`danger` 色彩，呼叫樣式可以覆寫。

獨立套件：`@docxkit/plugin-metric-card`，匯出 `metricCardPlugin`、`MetricCardOptions`、`MetricItem`。主套件 `docx-kit` 和 `docx-kit/node` 同樣匯出。

## 用法

```ts
import { createDocx, metricCardPlugin } from 'docx-kit'

const doc = createDocx().use(metricCardPlugin()).plugin('metricCard', {
  columns: 3,
  items: [
    { label: '營收', value: '128,000', unit: '元', trend: 'up', change: '+12%' },
    { label: '活躍使用者', value: 4200, trend: 'down', change: '-3%' },
    { label: '可用性', value: 99.9, unit: '%', trend: 'neutral', change: '持平' },
    { label: '事故', value: 0 },
  ],
})

await doc.toBlob()
```
