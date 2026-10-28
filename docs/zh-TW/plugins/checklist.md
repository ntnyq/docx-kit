# Checklist

產生驗收、巡檢、發佈清單，使用 Word 原生核取方塊。

## 選項

| 選項 | 型別 | 說明 |
| --- | --- | --- |
| `items` | `ChecklistItem[]` | 按輸入順序呈現；空陣列顯示 `(no items)` |
| `title` | `string` | 可選標題 |

`ChecklistItem` 包含必填的 `text`，以及可選的 `checked`（預設 `false`）、`owner`（負責人）、`note`（下方備註）、`group`（分組標題）。連續相同分組只顯示一次標題，分組再次出現會重新顯示標題；不排序。無分組項目會重設目前分組。

支援核取方塊內容控制項的 Word 閱讀器可以切換狀態；瀏覽器預覽可能只顯示初始狀態。負責人和備註使用主題 `muted` 色彩，呼叫樣式可以覆寫預設樣式。

獨立套件：`@docxkit/plugin-checklist`，匯出 `checklistPlugin`、`ChecklistOptions`、`ChecklistItem`。主套件 `docx-kit` 和 `docx-kit/node` 同樣匯出。

## 用法

```ts
import { checklistPlugin, createDocx } from 'docx-kit'

const doc = createDocx().use(checklistPlugin()).plugin('checklist', {
  title: '發佈檢查',
  items: [
    { text: '程式碼審查', checked: true, owner: 'Alice', group: '發佈前' },
    { text: '驗收測試', owner: 'Bob', note: '附上測試報告', group: '發佈前' },
    { text: '監控錯誤', group: '發佈後' },
  ],
})

await doc.toBlob()
```
