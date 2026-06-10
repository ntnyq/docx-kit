# CSS-like Styling

docx-kit uses CSS-like property names for styling, making document styling intuitive for web developers.

## Basic Styling

Style rules use familiar CSS property names that map to Word OOXML internally:

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

### All Style Properties

| Property | Type | Description |
|---|---|---|
| `fontSize` | `UnitValue` | Font size (bare number = pt) |
| `fontWeight` | `FontWeight` | Font weight: `'bold'`, `'normal'`, or `100`–`900` |
| `fontStyle` | `'italic' \| 'normal'` | Italic toggle |
| `fontFamily` | `LiteralUnion<'Arial' \| 'Calibri' \| 'Times New Roman'>` | Font family |
| `color` | `string \| HexColor` | Text / foreground color |
| `backgroundColor` | `string \| HexColor` | Background shading |
| `textAlign` | `'left' \| 'center' \| 'right' \| 'justify'` | Horizontal alignment |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | Vertical alignment (table cells) |
| `lineHeight` | `number \| UnitValue` | Line height |
| `letterSpacing` | `UnitValue` | Character spacing |
| `textIndent` | `UnitValue` | First-line indent |
| `underline` | `boolean \| 'single' \| 'double'` | Underline |
| `strike` | `boolean` | Strikethrough |
| `allCaps` | `boolean` | Force uppercase |
| `margin` | `UnitValue \| shorthand string` | CSS-like margin shorthand |
| `marginTop/Bottom/Left/Right` | `UnitValue` | Individual margins |
| `border` | `BorderRule` | All-sides border shorthand |
| `borderTop/Bottom/Left/Right` | `BorderRule` | Individual borders |
| `width` | `UnitValue` | Element width |
| `height` | `UnitValue` | Element height |
| `docx` | `Record<string, unknown>` | Raw passthrough to `docx` constructor |

## Unit Values

Length values can be expressed as bare numbers or explicit unit strings:

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

### Unit Conventions

| Context | Bare Number Meaning |
|---|---|
| `fontSize` | pt |
| `margin*` / `padding*` | pt |
| `textIndent` | pt |
| `letterSpacing` | pt |
| image `width` / `height` | px |

## Named Stylesheets (defineStyles)

Define reusable styles with `defineStyles()`:

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

Use them via `className` on nodes:

```ts
doc
  .h1('Title', { className: 'title' })
  .p('Body text', { className: 'body' })
  .p('Disclaimer', { className: 'small' })
  .p('Important!', { className: 'red' })
```

Multiple classes (space-separated or array):

```ts
doc.p('Important note', {
  className: 'body red',  // space-separated
})
```

## Style Priority (Cascade)

Styles are merged with the following priority (highest wins):

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

## Borders

Define borders per side with style, width, and color:

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

Available border styles: `'single'`, `'double'`, `'dashed'`, `'dotted'`, `'none'`.

## Margin Shorthand

CSS-like margin shorthand with 1, 2, or 4 values:

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

## Escape Hatch (docx)

For OOXML properties not yet covered by the CSS-like mapping, use `docx`:

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

## Defaults Per Element Type

Configure default styles for each element type in `createDocx()` config:

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

## Theme Tokens

Define reusable design tokens via the `theme` option:

```ts
const doc = createDocx({
  theme: {
    colors: {
      primary: '#2563eb',
      secondary: '#64748b',
      danger: '#dc2626',
    },
    fontFamily: {
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
      tight: '5pt',
      normal: '10pt',
      loose: '20pt',
    },
  },
})
```
