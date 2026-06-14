# Data Table

把物件陣列渲染成表格，支援欄位標題映射、格式化與對齊。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `data` | `Record<string, unknown>[]` | 資料陣列 |
| `labels` | `Record<string, string>` | 欄位標題映射 |
| `format` | `Record<string, 'currency' \| 'date' \| 'number' \| 'percent'>` | 欄位格式化 |
| `align` | `Record<string, 'left' \| 'center' \| 'right'>` | 欄位對齊 |
| `striped` | `boolean` | 斑馬紋 |
| `bordered` | `boolean` | 是否顯示邊框 |

## 用法

```ts
createDocx().use(dataTablePlugin()).plugin('dataTable', {
  data: [{ name: 'Alice', revenue: 12000 }],
  labels: { name: '姓名', revenue: '收入' },
  format: { revenue: 'currency' },
})
```
