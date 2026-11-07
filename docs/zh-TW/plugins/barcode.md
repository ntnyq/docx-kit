# 條碼（Barcode）

條碼外掛使用 `bwip-js` 產生一維條碼圖片，
並作為 PNG 內容嵌入文件。

```ts
import {
  barcodePlugin,
  type BarcodeFormat,
  type BarcodeOptions,
} from 'docx-kit'
```

## 安裝

使用 docx-kit 時同時安裝可選呈現器：

```shell
pnpm add bwip-js
```

## 選項

```ts
interface BarcodeOptions {
  text: string
  alignment?: 'center' | 'left' | 'right'
  backgroundColor?: string
  barColor?: string
  barHeight?: number
  caption?: string
  format?: BarcodeFormat
  includeText?: boolean
  rotate?: 'I' | 'L' | 'N' | 'R'
  scale?: number
  textColor?: string
  width?: number
}
```

常用格式包括 `code128`、`code39`、`code93`、`ean13`、`ean8`、
`interleaved2of5`、`isbn`、`itf14`、`upca` 與 `upce`。`format` 型別也
支援 `bwip-js` 提供的其他編碼器名稱。

## Code 128

```ts
import { barcodePlugin, createDocx } from 'docx-kit/node'

const doc = createDocx()
  .use(barcodePlugin())
  .plugin('barcode', {
    text: 'DOCX-KIT-2026',
    caption: 'Inventory identifier',
    format: 'code128',
    width: 280,
  })

await doc.save('barcode.docx')
```

## EAN-13

```ts
doc.plugin('barcode', {
  text: '5901234123457',
  format: 'ean13',
  barHeight: 16,
  includeText: true,
})
```

## 顏色與旋轉

```ts
doc.plugin('barcode', {
  text: 'DOCX-KIT-2026',
  backgroundColor: '#FEF3C7',
  barColor: '#1F2937',
  rotate: 'R',
  textColor: '#DC2626',
})
```

在 Node.js 中使用 PNG 呈現器，在瀏覽器建置中使用 canvas
呈現器。顯示高度根據產生的圖片計算，
因此改變 `width` 時會保持條碼寬高比。

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `BarcodeConfig` 和 `BarcodeStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `caption` | `TextStyleRule` | 文字排版 |
| `captionParagraph`, `imageParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

這些樣式控制 DOCX 段落與圖說；條碼像素、尺寸及圖片內文字繼續使用原有條碼選項。

```ts
barcodePlugin({
  styles: { caption: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
