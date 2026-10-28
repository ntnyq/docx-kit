# 插圖（Figure）

組合行內圖片、可選題注與可選來源，使用核心圖片編譯器處理格式檢測、尺寸、替代文字和圖片解析。

## 匯入

```ts
import { figurePlugin, type FigureOptions } from 'docx-kit'
```

獨立套件 `@docxkit/plugin-figure` 匯出相同工廠函式與型別。

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `data` | `string \| Uint8Array \| ArrayBuffer \| Blob` | 必填 | 圖片位元組、Blob、data URL、原始 base64 或已解析路徑 |
| `caption` | `string` | — | 圖片下方的原樣題注 |
| `source` | `string` | — | 題注下方的原樣來源，不添加前綴 |
| `alt` | `string` | — | 無障礙圖片描述 |
| `width` | `UnitValue` | 原始寬度 | 純數字按像素處理，支援 CSS 長度單位 |
| `height` | `UnitValue` | 原始高度 | 僅提供一個維度時保持寬高比 |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'` | 自動檢測 | 顯式圖片格式提示 |
| `alignment` | `'center' \| 'left' \| 'right'` | `'center'` | 圖片、題注與來源的對齊方式 |

同時提供寬高時直接使用指定尺寸。無法讀取原始尺寸時，核心圖片編譯器採用 300 × 180 像素的備援尺寸與寬高比。省略題注和來源不會產生額外段落。

## 範例

```ts
import { createDocx, figurePlugin } from 'docx-kit/node'

await createDocx()
  .use(figurePlugin())
  .plugin('figure', {
    data: './architecture.png',
    width: '12cm',
    alt: 'Services and their dependencies',
    caption: 'Figure 1. Service architecture',
    source: 'Source: Engineering team',
    alignment: 'center',
  })
  .save('architecture.docx')
```

瀏覽器中從 `docx-kit` 匯入，提供位元組、Blob 或 data URL，透過 `.toBlob()` 匯出。本地路徑需要 `docx-kit/node` 或顯式 `config.resolveImage` 適配器。URL 也需要顯式解析器，外掛不會自動獲取。詳見[圖片](/zh-TW/guide/images)。

圖片與題注段落在後面存在題注或來源時，會請求與下一段保持同頁。呼叫樣式可以覆蓋此行為和預設對齊方式。來源使用主題的 `muted` 顏色。

題注按輸入原樣呈現。自動編號、Word 序列域和交叉引用不在目前外掛範圍內。
