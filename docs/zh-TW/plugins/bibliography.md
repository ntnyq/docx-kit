# 參考文獻

`bibliographyPlugin` 產生數字編號文獻清單與可跳轉的引用段落。編號從 1 開始，按明確文獻目錄的順序決定；重複引用、提前引用與重複匯出維持相同編號。

```ts
import { bibliographyPlugin, type BibliographyEntry, type BibliographyOptions } from 'docx-kit'
// 獨立套件：import { bibliographyPlugin } from '@docxkit/plugin-bibliography'
```

獨立套件、`docx-kit`、`docx-kit/node` 均匯出工廠與 `BibliographyConfig`、`BibliographyOptions`、`BibliographyBaseOptions`、`BibliographyListOptions`、`BibliographyCitationOptions`、`BibliographyEntry`、`BibliographyStyles`。

## 選項

| 選項 | 型別 | 預設值與行為 |
| --- | --- | --- |
| `entries` | `BibliographyEntry[]` | 沿用工廠目錄，否則 `[]`；單次傳入會整體取代目錄，陣列順序決定編號 |
| `mode` | `'list' \| 'citation'` | 預設 `list`；文獻清單或引用段落 |
| `bookmarkPrefix` | `string` | 預設 `bib_`；清單書籤與引用目標的共同前綴 |
| `styles` | `BibliographyStyles` | 部位樣式 |
| `title` | `string` | 僅清單模式，預設 `References`；空字串省略標題 |
| `ids` | `string[]` | 引用模式必填；依傳入順序顯示，重複 ID 去重 |
| `hyperlink` | `boolean` | 僅引用模式，預設 `true`；編號連結至清單項目 |
| `prefix`、`suffix` | `string` | 僅引用模式，預設空字串；放在 `[1, 2]` 前後 |

工廠接受 `{ entries, styles }`。所有引用與對應清單須使用相同目錄及前綴；JSON/MCP 在各節點的 `options.entries` 中傳入目錄。清單模式空目錄、引用模式空 `ids` 均不輸出任何區塊，包括標題。`entries: []` 可清空工廠目錄。

`BibliographyEntry` 必填 `id`（唯一非空字串）與 `title`（非空白標題）；可選 `authors: string[]`、`year`、`publisher`、`doi`、`url`，後四項均為字串。`bookmarkPrefix + id` 須為 1–40 個 ASCII 字母、數字或底線，以字母開頭。

格式為 `[1] 作者. 標題. 年份. 出版者. DOI: 識別碼. URL`，缺少欄位省略，中繼資料按一般文字顯示。首版提供簡單數字格式，不包含 CSL、APA/GB/T 標準實作、Word 來源管理員或原生 `CITATION` / `BIBLIOGRAPHY` 功能變數；修改目錄後須重新產生文件。

引用輸出為區塊段落，透過 `prefix` / `suffix` 搭配內文，不是行內 DSL 節點。可跳轉引用需要文件中存在對應清單，清單可以稍後出現。每份清單轉譯一次；獨立清單使用不同 `bookmarkPrefix`，避免書籤重名。

缺少或型別不正確的引用 ID 陣列、未知引用 ID、重複來源 ID、無效書籤名稱與空白標題在轉譯時拋錯；編譯器包裝為 `PLUGIN_RENDER_FAILED`。

## 平台與樣式

支援 Node.js 與瀏覽器，僅需一般的 `@docxkit/core`、`docx` 對等依賴，無額外依賴。瀏覽器匯出用 `.toBlob()`，Node 檔案儲存從 `docx-kit/node` 呼叫 `.save()`。預覽顯示一般文字與內部連結，無須更新 Word 功能變數。

文字部位：`citation`、`entry`、`title`；段落部位：`citationParagraph`、`entryParagraph`、`titleParagraph`。標題使用主題 `primary`，文獻項目預設凸排。

工廠 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。優先順序：內建/主題預設 → 工廠樣式 → 通用節點樣式 → 單次部位樣式。支援單位與主題 token；文字、段落、儲存格分別使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整範例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-TW&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
