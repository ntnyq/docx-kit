# 交叉參照

`crossReferencePlugin` 產生參照段落，使用 Word 原生 `REF` / `PAGEREF` 功能變數，指向標題、圖說或其他文字上的現有書籤。

```ts
import { crossReferencePlugin, type CrossReferenceOptions } from 'docx-kit'
// 獨立套件：import { crossReferencePlugin } from '@docxkit/plugin-cross-reference'
```

獨立套件、`docx-kit`、`docx-kit/node` 均匯出工廠與 `CrossReferenceOptions`、`CrossReferenceConfig`、`CrossReferenceStyles`。

## 選項

| 選項 | 型別 | 預設值與行為 |
| --- | --- | --- |
| `target` | `string` | 必填；1–40 個 ASCII 字母、數字或底線，以字母開頭 |
| `kind` | `'text' \| 'number' \| 'page'` | 預設 `text`；分別使用 `REF`、帶 `\w` 完整編號內容的 `REF`、`PAGEREF` |
| `fallback` | `string` | 預設 `target`；更新功能變數之前顯示的快取文字 |
| `hyperlink` | `boolean` | 預設 `true`；加入 `\h` 跳至目標 |
| `prefix`、`suffix` | `string` | 預設空字串；功能變數前後的文字 |
| `styles` | `CrossReferenceStyles` | 部位樣式 |

用 `.bookmark('overview', ['概述'])` 建立目標，也可以在段落內放置書籤節點。`number` 要求書籤位於真正的編號段落中。一般圖說文字中的數字不是自動段落編號；這種情況可只為所需文字加入書籤，再用 `text` 參照。

外掛只產生參照，不建立或檢查目標是否存在。支援參照後面的書籤；同一文件中不要重複使用書籤名稱。空名稱、無效名稱與未知 `kind` 在轉譯時拋錯；目標不存在時，Word 更新功能變數會顯示參照錯誤。

## 平台與更新功能變數

支援 Node.js 與瀏覽器，僅需一般的 `@docxkit/core`、`docx` 對等依賴，無額外依賴。透過 `.use(crossReferencePlugin()).plugin('crossReference', options)` 使用；瀏覽器匯出用 `.toBlob()`，Node 檔案儲存從 `docx-kit/node` 呼叫 `.save()`。

輸出為區塊段落。瀏覽器預覽顯示 `fallback`，不計算頁碼或編號。需要在 Word 中更新功能變數以重新整理目標文字、編號與頁碼；結果取決於閱讀器支援。

## 樣式

文字部位：`text`；段落部位：`paragraph`。快取結果與前後文字都支援樣式。

工廠 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。優先順序：內建/主題預設 → 工廠樣式 → 通用節點樣式 → 單次部位樣式。支援單位與主題 token；文字、段落、儲存格分別使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整範例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-TW&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
