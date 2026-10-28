# 更新記錄（Changelog）

產生包含標題與四欄版本表格的發佈說明章節。

## 匯入

```ts
import { changelogPlugin, type ChangelogOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `entries` | `ChangelogEntry[]` | 必填 | 版本記錄清單 |
| `title` | `string` | `'Changelog'` | 表格上方的標題 |

`ChangelogEntry` 欄位：

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `version` | `string` | 版本標籤 |
| `date` | `string` | 發佈日期 |
| `type` | `'added' \| 'changed' \| 'fixed' \| 'removed'` | 帶有獨立高亮顏色的記錄分類 |
| `changes` | `string` | 可讀的變更摘要 |

## 範例

```ts
import { changelogPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(changelogPlugin())
  .plugin('changelog', {
    title: 'Release Notes',
    entries: [
      { version: '1.4.0', date: '2026-06-14', type: 'added', changes: 'Added zh-TW docs locale.' },
      { version: '1.4.1', date: '2026-06-15', type: 'fixed', changes: 'Corrected plugin sidebar links.' },
    ],
  })
```

## 說明

- `entries` 為空時仍呈現標題，並顯示居中的 `(no entries)` 預留位置內容。
- 適合發佈說明、遷移記錄或合規變更歷史。
