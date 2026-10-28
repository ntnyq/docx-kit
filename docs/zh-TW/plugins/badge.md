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
