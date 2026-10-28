# 樣式預設

樣式預設是為特定文件型別設定好的 `styles` + `defaults`（以及可選 `theme`）組合，可作為起點並按需覆蓋。

## 什麼是預設？

預設是一個凍結的 `DocxPreset` 物件，結構如下：

```ts
interface DocxPreset {
  id: string // e.g. 'modern'
  name: string // e.g. 'Modern'
  description: string
  config: DocxKitConfig // styles + defaults + theme
}
```

預設涵蓋 `DocxKitConfig` 的 `styles` 與 `defaults` 欄位。通常將預設的 `config` 展開傳入 `createDocx()`，再疊加自訂設定。

## 內建預設

docx-kit 提供 **3 個內建預設**：

| 預設 | ID | 適用場景 | 字型 |
| ---------------- | ---------- | ---------------------------------------- | --------------- |
| `classicPreset` | `classic` | 中文報告與公文信函 | SimHei / SimSun |
| `modernPreset` | `modern` | 商務報告、專業文件 | Calibri |
| `academicPreset` | `academic` | 學術論文、畢業論文 | Times New Roman |

也可以透過 `usePreset(id)` 按 ID 查找，或透過 `PRESET_LIST` 遍歷所有內建預設。

```ts
import {
  usePreset,
  PRESET_LIST,
  classicPreset,
  modernPreset,
  academicPreset,
} from 'docx-kit'

// All three are exported by name:
const modern = modernPreset
const classic = classicPreset
const academic = academicPreset

// Or look up by ID:
const found = usePreset('modern') // returns DocxPreset | undefined

// Or iterate all built-ins:
for (const p of PRESET_LIST) {
  console.log(p.id, p.name)
}
// classic Modern-style  academic
```

## 使用預設

最簡單的用法是將預設 `config` 展開傳入 `createDocx()`：

```ts
import { createDocx, modernPreset } from 'docx-kit/node'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')
  .p('Revenue grew 15% year-over-year.')
  .p('Key wins in the enterprise segment.')

await doc.save('report.docx')
```

文件會全文使用 Calibri、海軍藍與藍色標題、透過留白區隔內容、1.35 倍行距，以及垂直外邊距為 10pt 的居中圖片。

## 合併預設與自訂設定

先展開預設，再覆蓋指定欄位：

```ts
import { createDocx, modernPreset, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  styles: {
    ...modernPreset.config.styles,
    customClass: { color: '#f00', fontWeight: 'bold' },
  },
  metadata: {
    title: 'My Report',
    creator: 'Alice',
  },
})
```

## 預設詳解

### `classicPreset` — 公文樣式

適合報告、通知和公文信函的正式中文佈局。

```ts
import { createDocx, classicPreset } from 'docx-kit'

const doc = createDocx(classicPreset.config)

doc
  .h1('通知')
  .p('为深入贯彻落实……')
  .h2('一、总体要求')
  .p('以习近平新时代中国特色社会主义思想为指导……')
```

| 屬性 | 值 |
| --------- | -------------------------------- |
| `h1` 字型 | 黑體，22pt，加粗，居中 |
| `h2` 字型 | 黑體，16pt，加粗 |
| `h3` 字型 | 楷體，16pt，加粗 |
| 正文字型 | 宋體，14pt，1.5 倍行距 |
| 段落 | 首行縮進 28pt（2 個字符） |
| 圖片 | 居中，垂直外邊距 8pt |

### `modernPreset` — 商務樣式

簡潔專業，採用藍色強調色的商務樣式。

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 2026 Report')
  .h2('Executive Summary')
  .p('Revenue exceeded projections by 12%.')
```

| 屬性 | 值 |
| --------- | ---------------------------------------- |
| `h1` 字型 | Calibri，24pt，加粗，主題主色 |
| `h2` 字型 | Calibri，16pt，加粗，主題強調色 |
| `h3` 字型 | Calibri，13pt，加粗，主題強調色 |
| 正文字型 | Calibri，11pt，1.35 倍行距 |
| 圖片 | 居中，垂直外邊距 10pt |

### `academicPreset` — 學術與論文樣式

正式學術樣式，採用雙倍行距與兩端對齊。

```ts
import { createDocx, academicPreset } from 'docx-kit'

const doc = createDocx(academicPreset.config)

doc
  .h1('Abstract')
  .p('This paper proposes a novel approach to …')
  .h2('1. Introduction')
  .p('The field of …')
```

| 屬性 | 值 |
| --------- | --------------------------------------------------- |
| `h1` 字型 | Times New Roman，16pt，加粗，居中 |
| `h2` 字型 | Times New Roman，14pt，加粗 |
| `h4` 字型 | Times New Roman，12pt，加粗與斜體 |
| 正文字型 | Times New Roman，12pt，**雙倍行距**，兩端對齊 |
| 段落 | 首行縮進 24pt（2 個字符） |
| 圖片 | 居中，垂直外邊距 12pt |

## 結合主題

預設和主題互相獨立。預設定義_樣式規則_（h1 顏色、字級等），主題定義_設計令牌_（調色板、字型棧、尺寸體系），樣式透過 `$category.key` 引用這些令牌。

```ts
import { createDocx, modernPreset, useTheme } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  // theme tokens (e.g. $colors.primary) override literal colors when
  // a style rule references them, but presets don't currently
  // reference tokens — so use theme with custom styles for full effect
})
```

詳見[主題](/zh-TW/guide/themes)。

## 後續閱讀

- [主題](/zh-TW/guide/themes) — 將預設與主題令牌配合使用
- [樣式](/zh-TW/guide/styling) — 掌握樣式系統
- [範例：現代預設](/zh-TW/examples/preset-modern) — 完整商務報告範例
- [範例：學術預設](/zh-TW/examples/preset-academic) — 完整論文樣式範例

## 頁面佈局與 Word 編輯

所有內建預設均使用 A4。Modern 的上下邊距為 22 mm、左右為 24 mm；Classic 的上下邊距為 25 mm、左右為 28 mm；Academic 的四周邊距為 25.4 mm。標題與後續段落儘量保持同頁，並重置首行縮進。正文縮進不會影響外掛表格或題注。

原生 Word 文件預設值和標題樣式可在編輯匯出檔案時保留排版。拉丁字型與東亞字型分別由 `fontFamily` 和 `fontFamilyEastAsia` 設定。打開文件的機器需要安裝相應字型，否則 Word 會替換字型。
