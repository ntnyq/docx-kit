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

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `TocConfig` 和 `TocStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `title` | `TextStyleRule` | 文字排版 |
| `titleParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

此處僅自訂明確產生的標題。更新目錄欄位時，Word 使用文件的 TOC 樣式產生目錄項目。

```ts
tocPlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
