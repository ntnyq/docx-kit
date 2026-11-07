# 浮水印（Watermark）

透過帶樣式的段落呈現大號半透明文字浮水印，適合放在節的頁眉或頁腳。

## 匯入

```ts
import { watermarkPlugin, type WatermarkOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `text` | `string` | _必填_ | 浮水印文字 |
| `alignment` | `AlignmentType` | `CENTER` | 水平對齊 |
| `color` | `string` | `'BFBFBF'` | RRGGBB 十六進制文字顏色 |
| `fontSize` | `number` | `48` | 字級，單位為半磅（48 = 24pt） |

## 範例

### 基本浮水印

```ts
import { createDocx, watermarkPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Confidential Document')
  .p('This document contains proprietary information.')
  .plugin('watermark', { text: 'CONFIDENTIAL' })
  .save('confidential.docx')
```

### 紅色 DRAFT 浮水印

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Work in Progress')
  .p('This document is still under review.')
  .plugin('watermark', {
    text: 'DRAFT',
    color: 'FF0000',
    fontSize: 72, // 36pt
  })
  .save('draft.docx')
```

### 左對齊浮水印

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Internal Memo')
  .plugin('watermark', {
    text: 'INTERNAL USE ONLY',
    alignment: AlignmentType.LEFT,
    color: '999999',
  })
  .save('internal.docx')
```

### 放入節的頁眉或頁腳

要為整份文件添加浮水印，請放入頁眉或頁腳：

```ts
const doc = createDocx()
  .use(watermarkPlugin())

// Render watermark in a footer that applies to all pages
const wm = doc.plugin('watermark', {
  text: 'SAMPLE',
  color: 'D0D0D0',
  fontSize: 60,
})

// Note: The exact header/footer integration depends on
// how sections are configured in your document.
```

### 多個浮水印

```ts
const doc = createDocx()
  .use(watermarkPlugin())
  .h1('Review Document')
  .plugin('watermark', { text: 'DRAFT', color: 'FF6600' })
  .p('Please review the following content carefully.')
  .pageBreak()
  .plugin('watermark', { text: 'CONFIDENTIAL', color: 'CC0000' })
  .p('This section contains confidential financial data.')
  .save('multi-watermark.docx')
```

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `WatermarkConfig` 和 `WatermarkStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `text` | `TextStyleRule` | 文字排版 |
| `paragraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

```ts
watermarkPlugin({
  styles: { text: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
