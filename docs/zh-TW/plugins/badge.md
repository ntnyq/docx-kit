# 徽章（Badge）

呈現 `DRAFT`、`APPROVED`、`INTERNAL` 等簡短狀態標籤。

## 匯入

```ts
import { badgePlugin, type BadgeOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `text` | `string` | 必填 | 徽章文字 |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | `'neutral'` | 預設色調名稱或自訂文字顏色 |
| `backgroundColor` | `string` | 由預設決定 | 自訂徽章背景色 |

## 範例

```ts
import { badgePlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(badgePlugin())
  .plugin('badge', {
    text: 'DRAFT',
    color: 'warning',
  })
```

## 說明

- 預設提供配套的前景色與背景色。
- 自訂 `color` 若不屬於預設名稱，會按字面文字顏色處理。

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `BadgeConfig` 和 `BadgeStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `text` | `TextStyleRule` | 文字排版 |
| `paragraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

```ts
badgePlugin({
  styles: { text: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
