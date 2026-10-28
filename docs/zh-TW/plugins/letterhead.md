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
