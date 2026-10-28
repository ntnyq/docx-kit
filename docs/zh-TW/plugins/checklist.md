# Checklist

產生驗收、巡檢、發佈清單，預設使用 GitHub 風格的圖形核取方塊。

## 選項

| 選項 | 型別 | 說明 |
| --- | --- | --- |
| `items` | `ChecklistItem[]` | 按輸入順序呈現；空陣列顯示 `(no items)` |
| `checkboxStyle` | `'github' \| 'native'` | 預設 `'github'`，顯示靜態圓角圖形；`'native'` 使用可切換狀態的 Word 原生控制項 |
| `title` | `string` | 可選標題 |

`ChecklistItem` 包含必填的 `text`，以及可選的 `checked`（預設 `false`）、`owner`（負責人）、`note`（下方備註）、`group`（分組標題）。連續相同分組只顯示一次標題，分組再次出現會重新顯示標題；不排序。無分組項目會重設目前分組。

分組標題預設粗體，使用主題的 `accent` 色彩（退回 `primary`），字級約為正文的 120%，組間留白隨正文字級縮放。選填的清單總標題使用 `primary`，字級約為正文的 140%，保持高於分組的層級。呼叫時明確設定的文字色彩可覆寫這些預設色彩。

預設的 `'github'` 樣式使用 SVG 圖形，並提供高解析度 PNG 備援。勾選方塊使用主題的 `accent` 色彩（未設定時退回 `primary`），搭配白色勾號；未勾選方塊的邊線使用 `muted`（退回 `border`），底色使用 `background`。兩種圖片格式使用相同的主題色彩。圖形核取方塊會隨工作的實際字級縮放，並與文字視覺置中對齊；換行文字與備註保持相同的文字起點。這些是靜態圖片，不依賴符號字型。

設定 `checkboxStyle: 'native'` 後，支援核取方塊內容控制項的 Word 閱讀器可以切換狀態；瀏覽器預覽可能只顯示初始狀態。負責人和備註使用主題 `muted` 色彩，呼叫樣式可以覆寫預設文字與段落樣式。備註字級隨工作正文縮放，約為正文的 80%，依 Word 的半點精度取整。負責人字級約為正文的 85%，即使正文明確設定了色彩，負責人仍保留主題的 `muted` 色彩。

獨立套件：`@docxkit/plugin-checklist`，匯出 `checklistPlugin`、`ChecklistOptions`、`ChecklistItem`。主套件 `docx-kit` 和 `docx-kit/node` 同樣匯出。

## 完整範例

[在 Playground 中開啟完整範例](https://docx-kit-playground.ntnyq.dev/?lang=zh-TW&example=checklist)，可直接修改程式碼、預覽效果並下載 DOCX。

這份 Ocean 主題範例包含發佈前後兩個分組、負責人、備註，以及已勾選與未勾選的工作。正文為 12pt，分組標題為 14.5pt，負責人為 10pt，備註為 9.5pt。

![包含分組標題、負責人與備註的完整 Checklist 效果](/examples/checklist.png)

<<< @/../apps/playground/src/constants/templates/checklist.ts#checklist-example
