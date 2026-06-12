# Style Types

CSS-like style properties and the stylesheet system.

## `DocxStyleRule`

The core CSS-like style descriptor. All properties use familiar CSS names.

```ts
interface DocxStyleRule {
  // ── Text ──
  fontSize?: UnitValue
  fontWeight?: FontWeight
  fontStyle?: 'italic' | 'normal'
  fontFamily?: LiteralUnion<'Arial' | 'Calibri' | 'Times New Roman'>
  color?: string | HexColor
  backgroundColor?: string | HexColor
  highlight?: HighlightColor
  underline?: boolean | 'single' | 'double' | 'dash' | 'dotDash'
  strike?: boolean
  allCaps?: boolean
  letterSpacing?: UnitValue

  // ── Spacing ──
  lineHeight?: number | UnitValue

  // ── Alignment ──
  textAlign?: TextAlign
  textIndent?: UnitValue
  verticalAlign?: VerticalAlign

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
  | 'bold'
  | 'normal'
  | 100 | 200 | 300 | 400 | 500 | 600 | 700 | 800 | 900
```

## `HighlightColor`

```ts
type HighlightColor =
  | 'black' | 'blue' | 'cyan' | 'darkBlue' | 'darkCyan'
  | 'darkGray' | 'darkGreen' | 'darkMagenta' | 'darkRed'
  | 'darkYellow' | 'green' | 'magenta' | 'none' | 'red'
  | 'white' | 'yellow'
```

## `TextAlign`

```ts
type TextAlign = 'left' | 'center' | 'right' | 'justify'
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

## StyleSheet & Inheritance

### `StyleSheet`

```ts
type StyleSheet = Record<string, DocxStyleRule>
```

### `StyleSheetEntry`

An entry in the stylesheet that may extend other styles.

```ts
interface StyleSheetEntry extends DocxStyleRule {
  /** Single parent or ordered list of parents (later wins on conflict). */
  extends?: string | string[]
}
```

### Style Inheritance via `extends`

Styles can extend other styles for CSS-like inheritance:

```ts
const styles = defineStyles({
  baseText: {
    fontSize: 11,
    fontFamily: 'Calibri',
    color: '#333',
  },
  body: {
    extends: 'baseText',       // inherits fontSize, fontFamily, color
    lineHeight: 1.5,             // override
  },
  h2: {
    extends: 'body',            // 3-level chain: h2 → body → baseText
    fontSize: 18,
    fontWeight: 'bold',
  },
  warning: {
    extends: ['baseText', 'callout'],  // multiple inheritance
    color: '#92400e',
  },
})
```

### Style Cascade

Style resolution follows this priority (highest wins):

```
1. Inline style (passed directly to builder methods)
2. ClassName(s) from the stylesheet (merged left → right)
3. Defaults (per-element defaults from config)
4. Theme tokens ($category.key → resolved value)
```

Class names can be single strings, space-separated strings, or arrays:

```ts
doc.p('Hello', { className: 'accent' })
doc.p('Hello', { className: 'accent highlight' })
doc.p('Hello', { className: ['accent', 'highlight'] })
```

## `defineStyles()`

Type-safe stylesheet factory with compile-time validation.

```ts
function defineStyles<T extends Record<string, StyleSheetEntry>>(
  styles: T & { [K in keyof T]: T[K] extends StyleSheetEntry ? T[K] : never },
): { [K in keyof T]: T[K] }

// Usage
const styles = defineStyles({
  accent: { color: '#2563eb', fontWeight: 'bold' },
  muted:  { color: '#6b7280', fontSize: 10 },
})
// → styles.accent is { color: '#2563eb'; fontWeight: 'bold' }
```

## Style Properties Reference

| Category | Properties |
|---|---|
| **Text** | `fontFamily`, `fontSize`, `fontWeight`, `fontStyle`, `color`, `highlight`, `underline`, `strike`, `allCaps`, `backgroundColor`, `letterSpacing` |
| **Spacing** | `lineHeight` |
| **Margins** | `margin`, `marginTop`, `marginRight`, `marginBottom`, `marginLeft` |
| **Alignment** | `textAlign`, `textIndent`, `verticalAlign` |
| **Borders** | `border`, `borderTop`, `borderRight`, `borderBottom`, `borderLeft` |
| **Size** | `width`, `height` |
| **Escape** | `docx` — merge directly into docxjs options |
