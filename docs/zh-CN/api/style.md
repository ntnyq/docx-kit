# 样式类型

CSS 风格的样式属性与样式表系统。

## `DocxStyleRule`

核心 CSS 风格样式描述符，所有属性都使用熟悉的 CSS 名称。

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

## 样式表与继承

### `StyleSheet`

```ts
type StyleSheet = Record<string, DocxStyleRule>
```

### `StyleSheetEntry`

样式表中可以继承其他样式的条目。

```ts
interface StyleSheetEntry extends DocxStyleRule {
  /**
   * Single parent or ordered list of parents (later wins on conflict).
   */
  extends?: string | string[]
}
```

### 通过 `extends` 继承样式

样式可以继承其他样式，实现类似 CSS 的继承：

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

### 样式层叠

样式按以下优先级解析（高优先级覆盖低优先级）：

```
1. Inline style (passed directly to builder methods)
2. ClassName(s) from the stylesheet (merged left → right)
3. Defaults (per-element defaults from config)
4. Theme tokens ($category.key → resolved value)
```

类名可以是单个字符串、空格分隔的字符串或数组：

```ts
doc.p('Hello', { className: 'accent' })
doc.p('Hello', { className: 'accent highlight' })
doc.p('Hello', { className: ['accent', 'highlight'] })
```

## `defineStyles()`

支持编译时校验的类型安全样式表工厂函数。

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

## 样式属性参考

| 分类 | 属性 |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **文本** | `fontFamily`, `fontFamilyEastAsia`, `fontSize`, `fontWeight`, `fontStyle`, `bold`, `italic`, `color`, `highlight`, `underline`, `strike`, `doubleStrike`, `allCaps`, `smallCaps`, `subScript`, `superScript`, `rightToLeft`, `emboss`, `imprint`, `backgroundColor`, `characterSpacing`, `letterSpacing` |
| **间距** | `lineHeight`, `keepLines`, `keepNext`, `pageBreakBefore`, `widowControl` |
| **外边距** | `margin`, `marginTop`, `marginRight`, `marginBottom`, `marginLeft` |
| **对齐** | `textAlign`, `textIndent`, `verticalAlign`, `outlineLevel`, `tabStops` |
| **边框** | `border`, `borderTop`, `borderRight`, `borderBottom`, `borderLeft` |
| **尺寸** | `width`, `height` |
| **底层配置** | `docx` — 直接合并到 docxjs 选项 |

## `normalizeHexColor()`

将颜色传给原生 `docx` 选项前使用此工具规范化。核心样式和内置插件采用相同转换：三位十六进制色值展开为六位，并移除开头的 `#`；其他值保留原生校验行为。

```ts
import { normalizeHexColor } from 'docx-kit'

normalizeHexColor('#f00') // 'ff0000'
normalizeHexColor('#123456') // '123456'
```
