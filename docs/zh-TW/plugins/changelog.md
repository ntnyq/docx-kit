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

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `ChangelogConfig` 和 `ChangelogStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `bodyText`, `emptyText`, `headerText`, `title` | `TextStyleRule` | 文字排版 |
| `bodyParagraph`, `emptyParagraph`, `headerParagraph`, `titleParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |
| `bodyCell`, `headerCell` | `Omit<CellStyleRule, 'height'>` | 儲存格背景、框線、寬度及內距 |

```ts
changelogPlugin({
  styles: { headerText: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
