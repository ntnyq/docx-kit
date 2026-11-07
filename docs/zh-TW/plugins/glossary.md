# 術語與縮寫表

`glossaryPlugin` 產生術語、全稱、解釋三欄表格，支援排序與分組。

```ts
import { glossaryPlugin, type GlossaryEntry, type GlossaryOptions } from 'docx-kit'
// 獨立套件：import { glossaryPlugin } from '@docxkit/plugin-glossary'
```

獨立套件、`docx-kit`、`docx-kit/node` 均匯出工廠與 `GlossaryOptions`、`GlossaryEntry`、`GlossaryConfig`、`GlossaryStyles`。

## 選項

| 選項 | 型別 | 預設值與行為 |
| --- | --- | --- |
| `entries` | `GlossaryEntry[]` | 必填；空陣列不輸出任何區塊，包括標題 |
| `title` | `string` | 預設 `Glossary`；空字串省略標題 |
| `sort` | `'term' \| 'none'` | 預設 `term`，不區分大小寫排序術語與組名；`none` 保留首次出現順序 |
| `locale` | `string` | 預設 `en`；傳給 `Intl.Collator` 的排序語言 |
| `groupBy` | `'group' \| 'none'` | 預設 `none`；依組分別產生標題與表格 |
| `labels` | `{ term?, fullForm?, definition? }` | 字串欄名，預設 `Term`、`Full form`、`Definition` |
| `styles` | `GlossaryStyles` | 部位樣式 |

`GlossaryEntry` 必填 `term`（術語或縮寫）與 `definition`（解釋），兩項均不得為空白；可選 `fullForm`（全稱）與 `group`（組名），均為字串。沒有全稱時保留空白儲存格。

輸入陣列不會被修改。重複術語保留為獨立列；相同排序結果維持輸入順序。啟用分組後，不相鄰的同組項目也會合併。預設排序時無組名的項目排在最前，且不顯示組標題；`sort: 'none'` 時組按首次出現順序，組內項目維持原序。表頭跨頁重複，長列允許跨頁。

空白術語或解釋、無效 `sort` / `groupBy` 在轉譯時拋錯；無效排序語言由 `Intl.Collator` 拋出 `RangeError`。編譯器包裝為 `PLUGIN_RENDER_FAILED`。

## 平台與樣式

支援 Node.js 與瀏覽器，僅需一般的 `@docxkit/core`、`docx` 對等依賴，無額外依賴。瀏覽器匯出用 `.toBlob()`，Node 檔案儲存從 `docx-kit/node` 呼叫 `.save()`。預覽支援標題與表格，最終分頁由閱讀器決定。

文字部位：`definition`、`fullForm`、`group`、`headerText`、`term`、`title`；段落部位：`bodyParagraph`、`groupParagraph`、`headerParagraph`、`titleParagraph`；儲存格部位：`bodyCell`、`headerCell`。標題與組標題使用主題 `primary`，表頭底色使用 `surface`，框線使用 `border`。

工廠 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。優先順序：內建/主題預設 → 工廠樣式 → 通用節點樣式 → 單次部位樣式。支援單位與主題 token；文字、段落、儲存格分別使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整範例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-TW&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
