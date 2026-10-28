# 分隔線（Divider）

在章節之間插入橫線，或從視覺上分隔內容區塊。

## 匯入

```ts
import { dividerPlugin, type DividerOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | `'solid'` | 分隔線邊框樣式 |
| `color` | `string` | `'D9D9D9'` | 十六進制分隔線顏色 |
| `spacingBefore` | `number` | `200` | 分隔線段前間距 |
| `spacingAfter` | `number` | `200` | 分隔線段後間距 |

## 範例

```ts
import { createDocx, dividerPlugin } from 'docx-kit'

const doc = createDocx()
  .use(dividerPlugin())
  .h1('Quarterly Review')
  .plugin('divider', { style: 'double', color: '4472C4' })
  .p('Executive summary starts here.')
```
