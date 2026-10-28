# 目錄（Table of Contents）

插入根據標題層級產生的 Word 目錄域。

## 匯入

```ts
import { tocPlugin, type TocOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `title` | `string` | `'Contents'` | 可見章節標題 |
| `maxLevel` | `number` | `3` | 納入目錄的最大標題層級，限制在 `1..9` |

## 範例

```ts
import { createDocx, tocPlugin } from 'docx-kit'

const doc = createDocx()
  .use(tocPlugin())
  .plugin('toc', { title: 'Document Contents', maxLevel: 4 })
```

## 說明

- 外掛輸出真正的 DOCX 目錄域，而非靜態文字快照。
- Microsoft Word 打開檔案後可能提示更新域。
