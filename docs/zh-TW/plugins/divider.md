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

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `DividerConfig` 和 `DividerStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `paragraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

```ts
dividerPlugin({
  styles: { paragraph: { borderBottom: { color: '#2563EB', width: 2 }, marginTop: 0 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
