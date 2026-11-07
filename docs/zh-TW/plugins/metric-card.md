# 指標卡（Metric Card）

以等寬表格網格呈現報告指標。每張卡包含標籤、醒目的數值，以及可選單位和對比文字。

## 匯入

```ts
import { metricCardPlugin, type MetricCardOptions, type MetricItem } from 'docx-kit'
```

獨立套件 `@docxkit/plugin-metric-card` 匯出相同工廠函式與型別，呼叫名稱為 `metricCard`。

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `items` | `MetricItem[]` | 必填 | 按顯示順序排列的指標；空陣列呈現 `(no metrics)` |
| `columns` | `number` | `3` | 每行最多卡片數，必須為正安全整數 |

### `MetricItem`

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `label` | `string` | 指標名稱 |
| `value` | `number \| string` | 顯示值，包括零；不自動按地區格式化 |
| `unit` | `string` | 在空格後追加的單位 |
| `change` | `string` | 原樣輸出的對比文字，如 `+12% vs last month` |
| `trend` | `'up' \| 'down' \| 'neutral'` | 可選箭頭與顏色：↑ 成功色，↓ 危險色，→ 弱化色 |

`change` 和 `trend` 獨立工作，不會根據數值或變化文字推斷趨勢。下降代表好轉時，可以省略 `trend` 並提供自己的對比文字，或覆蓋呼叫顏色。

## 範例

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

每行不會跨頁拆分。最後一行補齊空儲存格；條目少於 `columns` 時，現有卡片平分全部寬度。無效欄數會拋出代碼為 `PLUGIN_RENDER_FAILED` 的 `DocxKitError`，空輸入也會檢查欄數。

網格使用主題的 `surface`、`primary`、`muted`、`success` 和 `danger` 顏色，並提供內建備援值。呼叫樣式可覆蓋文字、段落、儲存格背景與邊框。

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `MetricCardConfig` 和 `MetricCardStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `change`, `emptyText`, `label`, `unit`, `value` | `TextStyleRule` | 文字排版 |
| `changeParagraph`, `emptyParagraph`, `labelParagraph`, `paddingParagraph`, `valueParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |
| `cell` | `Omit<CellStyleRule, 'height'>` | 儲存格背景、框線、寬度及內距 |

```ts
metricCardPlugin({
  styles: { value: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
