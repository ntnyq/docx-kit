# 信頭（Letterhead）

為正式信函呈現居中的公司頁頭。

## 匯入

```ts
import { letterheadPlugin, type LetterheadOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `companyName` | `string` | 必填 | 主要品牌或公司名稱 |
| `tagline` | `string` | — | 名稱下方的可選副標題 |
| `phone` | `string` | — | 聯繫資訊行中的電話 |
| `email` | `string` | — | 聯繫資訊行中的電子郵件 |
| `website` | `string` | — | 聯繫資訊行中的網站 |
| `address` | `string` | — | 橫線上方的最後一行地址 |

## 範例

```ts
import { createDocx, letterheadPlugin } from 'docx-kit'

const doc = createDocx()
  .use(letterheadPlugin())
  .plugin('letterhead', {
    companyName: 'Acme Corp',
    tagline: 'Designing reliable document systems',
    phone: '+1 555 0100',
    email: 'info@acme.com',
    website: 'acme.com',
    address: '1 Market Street, San Francisco, CA',
  })
```

## 說明

- 外掛以藍色下邊框結束，正文可緊接其下開始。
- 聯繫資訊自動合併為一行居中文字。

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `LetterheadConfig` 和 `LetterheadStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `address`, `companyName`, `contact`, `tagline` | `TextStyleRule` | 文字排版 |
| `addressParagraph`, `companyParagraph`, `contactParagraph`, `ruleParagraph`, `taglineParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

```ts
letterheadPlugin({
  styles: { companyName: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
