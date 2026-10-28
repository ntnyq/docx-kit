# Figure

組合圖片、圖說和來源，重用核心圖片編譯器。

## 選項

| 選項 | 型別 | 說明 |
| --- | --- | --- |
| `data` | `string \| Uint8Array \| ArrayBuffer \| Blob` | 必填；支援位元組、Blob、data URL、base64 或已設定解析器的路徑 |
| `caption` | `string` | 可選圖說，按原文輸出 |
| `source` | `string` | 可選來源，顯示在圖說下方，不自動添加前綴 |
| `alt` | `string` | 圖片替代文字 |
| `width` / `height` | `UnitValue` | 數字以像素處理；支援 CSS 長度單位，只傳一邊會保持比例 |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'` | 預設自動偵測 |
| `alignment` | `'center' \| 'left' \| 'right'` | 圖片、圖說與來源對齊，預設置中 |

未傳入尺寸時使用圖片原始尺寸；無法讀取時退回 300 × 180 像素及其比例。兩邊都傳入時使用指定尺寸。圖說和來源均為可選，不產生多餘空段落。

瀏覽器從 `docx-kit` 匯入，傳入位元組、Blob 或 data URL，並用 `.toBlob()` 匯出。本機路徑需要 `docx-kit/node` 或 `config.resolveImage`，URL 也需要明確解析器，不會自動下載。圖片與後續圖說、來源預設盡量保持同頁，呼叫樣式可以覆寫對齊等設定。來源使用主題 `muted` 色彩。

首版不產生自動編號、Word 序列欄位或交互參照，需要的編號請寫入 `caption`。

獨立套件：`@docxkit/plugin-figure`，匯出 `figurePlugin`、`FigureOptions`。主套件 `docx-kit` 和 `docx-kit/node` 同樣匯出。

## 用法

```ts
import { createDocx, figurePlugin } from 'docx-kit/node'

await createDocx().use(figurePlugin()).plugin('figure', {
  data: './architecture.png',
  width: '12cm',
  alt: '服務及其相依關係',
  caption: '圖 1. 服務架構',
  source: '來源：工程團隊',
}).save('architecture.docx')
```
