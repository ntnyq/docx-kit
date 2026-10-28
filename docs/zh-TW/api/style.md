# 樣式型別

CSS 風格的樣式屬性與樣式表系統。

## `DocxStyleRule`

核心 CSS 風格樣式描述符，所有屬性都使用熟悉的 CSS 名稱。

```ts
interface DocxStyleRule {
  // ── Text ──
  fontSize?: UnitValue
  fontWeight?: FontWeight
  fontStyle?: 'italic' | 'normal'
  bold?: boolean
  italic?: boolean
  fontFamily?: LiteralUnion<'Arial' | 'Calibri' | 'Times New Roman'>
  fontFamilyEastAsia?: string // independent Chinese, Japanese, and Korean font
  color?: string | HexColor
  backgroundColor?: string | HexColor
  highlight?: HighlightColor
  underline?: boolean | 'single' | 'double'
  strike?: boolean
  doubleStrike?: boolean
  allCaps?: boolean
  smallCaps?: boolean
  subScript?: boolean
  superScript?: boolean
  rightToLeft?: boolean
  emboss?: boolean
  imprint?: boolean
  characterSpacing?: number
  letterSpacing?: UnitValue

  // ── Spacing ──
  lineHeight?: number | UnitValue
  keepLines?: boolean
  keepNext?: boolean
  pageBreakBefore?: boolean
  widowControl?: boolean

  // ── Alignment ──
  textAlign?: TextAlign
  textIndent?: UnitValue
  verticalAlign?: VerticalAlign
  outlineLevel?: number
  tabStops?: TabStopRule[]

  // ── Box model ──
  margin?:
    | UnitValue
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
  marginTop?: UnitValue
  marginBottom?: UnitValue
  marginLeft?: UnitValue
  marginRight?: UnitValue

  // ── Border ──
  border?: BorderRule
  borderTop?: BorderRule
  borderBottom?: BorderRule
  borderLeft?: BorderRule
  borderRight?: BorderRule

  // ── Size ──
  width?: UnitValue
  height?: UnitValue

  // ── Escape hatch ──
  docx?: Record<string, unknown>
}
```

## `FontWeight`

```ts
type FontWeight =
  'bold' | 'normal' | 100 | 200 | 300 | 400 | 500 | 600 | 700 | 800 | 900
```

## `HighlightColor`

```ts
type HighlightColor =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGray'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'green'
  | 'magenta'
  | 'none'
  | 'red'
  | 'white'
  | 'yellow'
```

## `TextAlign`

```ts
type TextAlign = 'left' | 'center' | 'right' | 'justify'
```

## `TabStopRule`

```ts
interface TabStopRule {
  position: UnitValue
  type:
    | 'bar'
    | 'center'
    | 'clear'
    | 'decimal'
    | 'end'
    | 'left'
    | 'num'
    | 'right'
    | 'start'
  leader?: 'dot' | 'hyphen' | 'middleDot' | 'none' | 'underscore'
}
```

## `VerticalAlign`

```ts
type VerticalAlign = 'top' | 'middle' | 'bottom'
```

## `BorderRule`

```ts
interface BorderRule {
  color?: string | HexColor
  style?: BorderStyle
  width?: UnitValue
}
```

## `BorderStyle`

```ts
type BorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'none'
```

## 樣式表與繼承

### `StyleSheet`

```ts
type StyleSheet = Record<string, DocxStyleRule>
```

### `StyleSheetEntry`

樣式表中可以繼承其他樣式的條目。

```ts
interface StyleSheetEntry extends DocxStyleRule {
  /**
   * Single parent or ordered list of parents (later wins on conflict).
   */
  extends?: string | string[]
}
```

### 透過 `extends` 繼承樣式

樣式可以繼承其他樣式，實現類似 CSS 的繼承：

```ts
const styles = defineStyles({
  baseText: {
    fontSize: 11,
    fontFamily: 'Calibri',
    color: '#333',
  },
  body: {
    extends: 'baseText', // inherits fontSize, fontFamily, color
    lineHeight: 1.5, // override
  },
  h2: {
    extends: 'body', // 3-level chain: h2 → body → baseText
    fontSize: 18,
    fontWeight: 'bold',
  },
  warning: {
    extends: ['baseText', 'callout'], // multiple inheritance
    color: '#92400e',
  },
})
```

### 樣式層疊

樣式按以下優先級解析（高優先級覆蓋低優先級）：

```
1. Inline style (passed directly to builder methods)
2. ClassName(s) from the stylesheet (merged left → right)
3. Defaults (per-element defaults from config)
4. Theme tokens ($category.key → resolved value)
```

類名可以是單個字串、空格分隔的字串或陣列：

```ts
doc.p('Hello', { className: 'accent' })
doc.p('Hello', { className: 'accent highlight' })
doc.p('Hello', { className: ['accent', 'highlight'] })
```

## `defineStyles()`

支援編譯時驗證的型別安全樣式表工廠函式。

```ts
function defineStyles<T extends Record<string, StyleSheetEntry>>(
  styles: T & { [K in keyof T]: T[K] extends StyleSheetEntry ? T[K] : never },
): { [K in keyof T]: T[K] }

// Usage
const styles = defineStyles({
  accent: { color: '#2563eb', fontWeight: 'bold' },
  muted: { color: '#6b7280', fontSize: 10 },
})
// → styles.accent is { color: '#2563eb'; fontWeight: 'bold' }
```

## 樣式屬性參考

| 分類 | 屬性 |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **文字** | `fontFamily`, `fontFamilyEastAsia`, `fontSize`, `fontWeight`, `fontStyle`, `bold`, `italic`, `color`, `highlight`, `underline`, `strike`, `doubleStrike`, `allCaps`, `smallCaps`, `subScript`, `superScript`, `rightToLeft`, `emboss`, `imprint`, `backgroundColor`, `characterSpacing`, `letterSpacing` |
| **間距** | `lineHeight`, `keepLines`, `keepNext`, `pageBreakBefore`, `widowControl` |
| **外邊距** | `margin`, `marginTop`, `marginRight`, `marginBottom`, `marginLeft` |
| **對齊** | `textAlign`, `textIndent`, `verticalAlign`, `outlineLevel`, `tabStops` |
| **邊框** | `border`, `borderTop`, `borderRight`, `borderBottom`, `borderLeft` |
| **尺寸** | `width`, `height` |
| **底層設定** | `docx` — 直接合併到 docxjs 選項 |
