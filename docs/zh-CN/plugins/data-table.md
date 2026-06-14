# Data Table

把对象数组渲染为表格，支持列标题映射、格式化和对齐。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `data` | `Record<string, unknown>[]` | 数据数组 |
| `labels` | `Record<string, string>` | 列标题映射 |
| `format` | `Record<string, 'currency' \| 'date' \| 'number' \| 'percent'>` | 列格式化 |
| `align` | `Record<string, 'left' \| 'center' \| 'right'>` | 列对齐 |
| `striped` | `boolean` | 斑马纹 |
| `bordered` | `boolean` | 是否显示边框 |

## 用法

```ts
createDocx().use(dataTablePlugin()).plugin('dataTable', {
  data: [{ name: 'Alice', revenue: 12000 }],
  labels: { name: '姓名', revenue: '收入' },
  format: { revenue: 'currency' },
})
```
