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
| `highlight` | `HighlightColor` | Text highlighting (`'yellow'`, `'green'`, `'red'`, etc.) |
| `textAlign` | `'left' \| 'center' \| 'right' \| 'justify'` | Horizontal alignment |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | Vertical alignment (table cells) |
| `lineHeight` | `number \| UnitValue` | Line height |
| `letterSpacing` | `UnitValue` | Character spacing |
| `textIndent` | `UnitValue` | First-line indent |
| `underline` | `boolean \| 'single' \| 'double' \| 'dash' \| 'dotDash'` | Underline |
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

Define reusable design tokens via the `theme` option, then reference them in styles using `$category.key` syntax:

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

### Built-in Themes

docx-kit ships with 3 built-in themes:

```ts
import { useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),   // 'minimal' | 'ocean' | 'warm'
})
```

| Theme | Palette | Best For |
|---|---|---|
| `minimal` | Clean grayscale + blue accent | General-purpose docs |
| `ocean` | Deep blue / teal color scheme | Business reports |
| `warm` | Warm earth-tone color scheme | Creative docs, presentations |

## Style Presets

For a faster start, docx-kit ships with **3 pre-configured style presets**. See [Style Presets](/guide/presets) for details.

Quick example using `modernPreset`:

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')           // Navy 26pt with blue underline (Calibri)
  .p('Body text here.')      // Calibri 11pt, 1.5× line height
```

## Themes

Themes are design tokens (colors, fonts, scale) that can be referenced in styles via `$category.key` syntax. docx-kit ships with **3 built-in themes** (`minimal`, `ocean`, `warm`). See [Themes](/guide/themes) for the full token reference.

Quick example using the ocean theme:

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

## See Also

- [Style Presets](/guide/presets) — Pre-configured style bundles
- [Themes](/guide/themes) — Design tokens
- [Builder API](/guide/builder-api) — Full builder reference
- [Examples: Theme System](/examples/theme-system) — End-to-end themed report

## Style Inheritance (extends)

Styles can extend other styles for CSS-like inheritance:

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

Inheritance is resolved at compile time with circular reference detection.

See the [Style Inheritance Example](/examples/style-inheritance) for a full demonstration.

## Style Presets

Quick-start with a pre-configured style preset:

```ts
import { createDocx, usePreset } from 'docx-kit'

// 3 built-in presets available:
const doc = createDocx(usePreset('modern').config)
// or 'classic' (gov-doc style), 'academic' (thesis style)
```

Each preset provides a complete stylesheet, page config, and element defaults tuned for its use case. You can also merge a preset with your own config:

```ts
const doc = createDocx({
  ...usePreset('modern').config,
  styles: {
    ...usePreset('modern').config.styles,
    customClass: { color: '#f00' },
  },
})
```
