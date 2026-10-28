# CSS 風格樣式

docx-kit 使用 CSS 風格屬性名，讓網頁開發者能直觀地設置文件樣式。

## 基本樣式

樣式規則使用熟悉的 CSS 屬性名，內部映射到 Word OOXML：

```ts
const style = {
  fontSize: 14,          // pt (bare number = pt)
  fontWeight: 'bold',    // 'bold' | 'normal' | 100–900
  fontStyle: 'italic',   // 'italic' | 'normal'
  fontFamily: 'Arial',   // font family
  color: '#333333',      // text color (hex or named)
  backgroundColor: '#fff3cd', // shading
  textAlign: 'center',   // 'left' | 'center' | 'right' | 'justify'
  lineHeight: 1.5,       // multiplier or UnitValue
  letterSpacing: 1,      // character spacing in pt
  underline: true,       // boolean | 'single' | 'double'
  strike: false,         // strikethrough
  allCaps: false,        // force uppercase
}
```

### 全部樣式屬性

| 屬性 | 型別 | 說明 |
|---|---|---|
| `fontSize` | `UnitValue` | 字級，純數字單位為 pt |
| `fontWeight` | `FontWeight` | 字重：`'bold'`、`'normal'` 或 `100`–`900` |
| `fontStyle` | `'italic' \| 'normal'` | 斜體開關 |
| `fontFamily` | `LiteralUnion<'Arial' \| 'Calibri' \| 'Times New Roman'>` | 字型；DOCX 使用 CSS 備援字型棧中的第一個字型 |
| `color` | `string \| HexColor` | 文字或前景顏色 |
| `backgroundColor` | `string \| HexColor` | 背景底紋 |
| `highlight` | `HighlightColor` | 文字高亮（`'yellow'`、`'green'`、`'red'` 等） |
| `textAlign` | `'left' \| 'center' \| 'right' \| 'justify'` | 水平對齊 |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | 垂直對齊（表格儲存格） |
| `lineHeight` | `number \| UnitValue` | 行高 |
| `letterSpacing` | `UnitValue` | 字符間距 |
| `textIndent` | `UnitValue` | 首行縮進 |
| `underline` | `boolean \| 'single' \| 'double' \| 'dash' \| 'dotDash'` | 下劃線 |
| `strike` | `boolean` | 刪除線 |
| `allCaps` | `boolean` | 強制大寫 |
| `margin` | `UnitValue \| 简写字符串` | CSS 風格外邊距簡寫 |
| `marginTop/Bottom/Left/Right` | `UnitValue` | 各方向外邊距 |
| `border` | `BorderRule` | 四邊邊框簡寫 |
| `borderTop/Bottom/Left/Right` | `BorderRule` | 各方向邊框 |
| `width` | `UnitValue` | 元素寬度 |
| `height` | `UnitValue` | 元素高度 |
| `docx` | `Record<string, unknown>` | 直接透傳給 `docx` 構造函式 |

## 單位值

長度可以使用純數字或帶顯式單位的字串：

```ts
const style = {
  fontSize: 14,            // bare number → pt
  marginLeft: '20px',      // explicit px
  marginTop: '1cm',        // explicit cm
  marginBottom: '10mm',    // explicit mm
  textIndent: '2em',       // ⚠️ not supported, use pt instead
  width: '50%',            // percentage (tables)
}
```

### 單位約定

| 場景 | 純數字的單位 |
|---|---|
| `fontSize` | pt |
| `margin*` / `padding*` | pt |
| `textIndent` | pt |
| `letterSpacing` | pt |
| 圖片 `width` / `height` | px |

## 命名樣式表（defineStyles）

透過 `defineStyles()` 定義可重用樣式：

```ts
import { defineStyles } from 'docx-kit'

const styles = defineStyles({
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#1a1a2e',
    textAlign: 'center',
    marginBottom: 10,
  },
  body: {
    fontSize: 12,
    lineHeight: 1.5,
    fontFamily: 'Calibri',
  },
  small: {
    fontSize: 9,
    color: '#999',
  },
  red: {
    color: '#dc3545',
    fontWeight: 'bold',
  },
})
```

透過節點的 `className` 使用：

```ts
doc
  .h1('Title', { className: 'title' })
  .p('Body text', { className: 'body' })
  .p('Disclaimer', { className: 'small' })
  .p('Important!', { className: 'red' })
```

多個類名（空格分隔或陣列）：

```ts
doc.p('Important note', {
  className: 'body red',  // space-separated
})
```

## 樣式優先級（層疊）

樣式按以下優先級合併，高優先級覆蓋低優先級：

```
inline style > className > element defaults
```

```ts
const doc = createDocx({
  styles: defineStyles({
    blue: { color: '#0000ff', fontSize: 14 },
  }),
  defaults: {
    text: { fontSize: 12, color: '#000' },
  },
})

// Result: fontSize=14 (from class), color=#ff0000 (inline wins over blue)
doc.p('This is red, 14pt', {
  className: 'blue',
  style: { color: '#ff0000' },
})
```

## 邊框

按方向設置邊框樣式、寬度和顏色：

```ts
doc.p('Bordered paragraph', {
  style: {
    border: {
      style: 'single',
      width: 1,
      color: '#333',
    },
  },
})

// Per-side overrides
doc.p('Mixed borders', {
  style: {
    border: { style: 'single', width: 1, color: '#ccc' },
    borderBottom: { style: 'double', width: 2, color: '#333' },
    borderLeft: { style: 'dashed', width: 1, color: '#999' },
  },
})
```

可用邊框樣式：`'single'`、`'double'`、`'dashed'`、`'dotted'`、`'none'`。

## 外邊距簡寫

與 CSS 類似，支援 1、2 或 4 個值的外邊距簡寫：

```ts
doc.p('Uniform margin', {
  style: { margin: '10pt' },
})

doc.p('Vertical | Horizontal', {
  style: { margin: '10pt 20pt' },
})

doc.p('Top Right Bottom Left', {
  style: { margin: '10pt 20pt 15pt 20pt' },
})
```

## 底層設定（docx）

對於尚未映射為 CSS 風格屬性的 OOXML 選項，可以使用 `docx`：

```ts
doc.p('Custom run properties', {
  style: {
    docx: {
      run: {
        emboss: true,     // embossed text effect
        shadow: true,     // shadow effect
      },
    },
  },
})
```

## 各元素型別的預設值

在 `createDocx()` 設定中為每種元素設置預設樣式：

```ts
const doc = createDocx({
  defaults: {
    text: {
      fontFamily: 'Calibri',
      fontSize: 11,
    },
    paragraph: {
      textAlign: 'justify',
      lineHeight: 1.5,
    },
    table: {
      bordered: true,
    },
    cell: {
      verticalAlign: 'middle',
      fontSize: 10,
    },
  },
})
```

## 主題令牌

透過 `theme` 選項定義可重用設計令牌，再用 `$category.key` 語法在樣式中引用：

```ts
import { createDocx, defineStyles } from 'docx-kit'

const doc = createDocx({
  theme: {
    colors: {
      primary: '#2563eb',
      secondary: '#64748b',
      danger: '#dc2626',
    },
    fonts: {
      heading: 'Arial',
      body: 'Calibri',
    },
    fontSize: {
      xs: 9,
      sm: 11,
      md: 12,
      lg: 16,
      xl: 20,
    },
    spacing: {
      tight: 5,
      normal: 10,
      loose: 20,
    },
  },
  styles: defineStyles({
    title: {
      color: '$colors.primary',        // resolves to '#2563eb'
      fontFamily: '$fonts.heading',    // resolves to 'Arial'
      fontSize: '$fontSize.xl',        // resolves to 20
      marginBottom: '$spacing.loose',  // resolves to 20
    },
    body: {
      fontSize: '$fontSize.md',
      fontFamily: '$fonts.body',
    },
  }),
})
```

### 內建主題

docx-kit 提供 3 個內建主題：

```ts
import { useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),   // 'minimal' | 'ocean' | 'warm'
})
```

| 主題 | 配色 | 適用場景 |
|---|---|---|
| `minimal` | 簡潔灰階與藍色強調色 | 通用文件 |
| `ocean` | 深藍與青綠色配色 | 商務報告 |
| `warm` | 暖色大地色系 | 創意文件、演示材料 |

## 樣式預設

為便於快速開始，docx-kit 提供 **3 個預設定樣式預設**。詳見[樣式預設](/zh-TW/guide/presets)。

使用 `modernPreset` 的簡短範例：

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')           // Navy 26pt with blue underline (Calibri)
  .p('Body text here.')      // Calibri 11pt, 1.5× line height
```

## 主題

主題是顏色、字型與尺寸體系等設計令牌，在樣式中透過 `$category.key` 引用。docx-kit 提供 **3 個內建主題**（`minimal`、`ocean`、`warm`），完整令牌參考見[主題](/zh-TW/guide/themes)。

使用 Ocean 主題的簡短範例：

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({
    title: {
      color: '$colors.primary',      // → '#0f172a'
      fontFamily: '$fonts.heading',  // → 'Georgia, serif'
      fontSize: '$fontSize.xl',      // → 20
      marginBottom: '$spacing.lg',   // → 28
    },
  }),
})
```

## 相關文件

- [樣式預設](/zh-TW/guide/presets) — 預設定樣式組合
- [主題](/zh-TW/guide/themes) — 設計令牌
- [Builder API](/zh-TW/guide/builder-api) — 完整建構器參考
- [範例：主題系統](/zh-TW/examples/theme-system) — 完整主題報告

## 樣式繼承（extends）

樣式可以繼承其他樣式，實現類似 CSS 的繼承：

```ts
const styles = defineStyles({
  // Base style
  baseText: {
    fontSize: 11,
    fontFamily: 'Calibri',
    color: '#333',
  },

  // inherit from baseText, override selectively
  body: {
    extends: 'baseText',
    lineHeight: 1.5,
  },

  // 3-level chain: h2 → body → baseText
  h2: {
    extends: 'body',
    fontSize: 18,
    fontWeight: 'bold',
  },

  // Multiple inheritance (later array entries win on conflicts)
  warning: {
    extends: ['baseText', 'callout'],
    color: '#92400e',
    border: { style: 'single', width: 1, color: '#f59e0b' },
  },
})
```

繼承在文件編譯階段解析，並檢測循環引用。

完整演示見[樣式繼承範例](/zh-TW/examples/style-inheritance)。

## 樣式預設

使用預設定樣式預設快速開始：

```ts
import { createDocx, usePreset } from 'docx-kit'

// 3 built-in presets available:
const doc = createDocx(usePreset('modern').config)
// or 'classic' (gov-doc style), 'academic' (thesis style)
```

每個預設都為對應場景提供完整樣式表、頁面設定與元素預設值。也可以將預設與自訂設定合併：

```ts
const doc = createDocx({
  ...usePreset('modern').config,
  styles: {
    ...usePreset('modern').config.styles,
    customClass: { color: '#f00' },
  },
})
```
