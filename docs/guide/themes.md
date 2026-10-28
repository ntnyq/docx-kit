# Themes

Themes are **semantic design tokens** — colors, fonts, font sizes, and spacing scales — that can be referenced from your style rules via the `$category.key` syntax. They decouple _what styles you define_ from _what colors and fonts you use_. Theme typography also supplies native Word defaults, so unstyled text and plugins inherit the document font automatically.

## What is a Theme?

A `DocxTheme` is a bag of design tokens:

```ts
interface DocxTheme {
  id?: string
  name?: string
  colors?: Record<string, string> // palette
  fonts?: Record<string, string> // font stacks
  fontSize?: Record<string, number> // scale (pt)
  spacing?: Record<string, number> // scale (pt)
}
```

**Why use a theme?**

- Switch the entire look of a document by changing one field
- Keep colors consistent across many documents
- Reuse token names like `primary`, `accent`, `success` semantically
- Pair with [presets](/guide/presets) for a complete starting config

## Built-in Themes

docx-kit ships with **3 built-in themes**:

| Theme          | ID        | Palette                             | Best For               | Fonts             |
| -------------- | --------- | ----------------------------------- | ---------------------- | ----------------- |
| `minimalTheme` | `minimal` | Clean grayscale + blue accent       | General business docs  | Calibri           |
| `oceanTheme`   | `ocean`   | Navy / teal with pale gray surfaces | Reports, articles      | Cambria / Calibri |
| `warmTheme`    | `warm`    | Warm gray / brown                   | Correspondence, essays | Georgia           |

Helpers:

```ts
import {
  useTheme,
  THEME_LIST,
  minimalTheme,
  oceanTheme,
  warmTheme,
} from 'docx-kit'

// All three are exported by name:
const ocean = oceanTheme
const minimal = minimalTheme
const warm = warmTheme

// Or look up by ID:
const found = useTheme('ocean') // returns DocxTheme | undefined

// Or iterate all built-ins:
for (const t of THEME_LIST) {
  console.log(t.id, t.name)
}
// minimal  ocean  warm
```

## Using a Theme

Pass the theme into `createDocx()`:

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({
    title: {
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      fontSize: '$fontSize.xl',
      marginBottom: '$spacing.lg',
    },
  }),
})

doc.h1('Annual Report 2026', { className: 'title' })
```

Any literal string starting with `$` in a style rule is treated as a token reference and resolved at compile time.

## Token Categories

### `$colors.*` — Palette

Each built-in theme provides 12 semantic color tokens:

| Token        | Purpose                | `minimal` | `ocean`   | `warm`    |
| ------------ | ---------------------- | --------- | --------- | --------- |
| `primary`    | Headings, primary text | `#111827` | `#0f172a` | `#44403c` |
| `secondary`  | Secondary text         | `#374151` | `#1e293b` | `#57534e` |
| `accent`     | Highlight / link color | `#2563eb` | `#0d9488` | `#8b5e34` |
| `success`    | Success state          | `#2e7d32` | `#14765c` | `#15803d` |
| `warning`    | Warning state          | `#9a6700` | `#946200` | `#a34c13` |
| `danger`     | Error state            | `#b91c1c` | `#e11d48` | `#b91c1c` |
| `info`       | Info state             | `#245a96` | `#24758a` | `#6b6578` |
| `muted`      | Muted text             | `#6b7280` | `#64748b` | `#78716c` |
| `background` | Page background        | `#ffffff` | `#ffffff` | `#ffffff` |
| `surface`    | Card / panel surface   | `#f9fafb` | `#f0f5f7` | `#f6f3ee` |
| `text`       | Body text              | `#1f2937` | `#1e293b` | `#292524` |
| `border`     | Border color           | `#e5e7eb` | `#cbd9df` | `#ddd6ca` |

### `$fonts.*` — Font Stacks

Each built-in theme provides body, heading, and code fonts, plus separate `bodyEastAsia` and `headingEastAsia` fonts for Chinese, Japanese, and Korean text. Word accepts one font per script; CSS fallback stacks are reduced to their first family.

| Token     | `minimal`  | `ocean`    | `warm`     |
| --------- | ---------- | ---------- | ---------- |
| `heading` | `Calibri`  | `Cambria`  | `Georgia`  |
| `body`    | `Calibri`  | `Calibri`  | `Georgia`  |
| `code`    | `Consolas` | `Consolas` | `Consolas` |

`bodyEastAsia` and `headingEastAsia` use Microsoft YaHei in Minimal and Ocean.
Warm pairs SimSun body text with SimHei headings. These are Word font names;
the application opening the document chooses a substitute when a font is absent.

### `$fontSize.*` — Type Scale

Each built-in theme provides a 5-step scale (values in pt):

| Token  | `minimal` | `ocean` | `warm` |
| ------ | --------- | ------- | ------ |
| `xs`   | 8         | 8       | 8      |
| `sm`   | 9         | 10      | 10     |
| `base` | 11        | 12      | 12     |
| `lg`   | 14        | 16      | 15     |
| `xl`   | 24        | 26      | 24     |

### `$spacing.*` — Spacing Scale

Values in pt. Use with `margin*`, `padding*`, `borderWidth` etc.:

| Token | `minimal` | `ocean` | `warm` |
| ----- | --------- | ------- | ------ |
| `xs`  | 4         | 5       | 4      |
| `sm`  | 8         | 10      | 9      |
| `md`  | 16        | 18      | 16     |
| `lg`  | 24        | 28      | 26     |
| `xl`  | 32        | 36      | 34     |

## Custom Theme

Define your own theme by passing a `DocxTheme` object to `createDocx()`:

```ts
import { createDocx, defineStyles, type DocxTheme } from 'docx-kit'

const brand: DocxTheme = {
  id: 'brand-2026',
  name: 'Brand 2026',
  colors: {
    primary: '#1a56db',
    accent: '#f59e0b',
    danger: '#dc2626',
    success: '#16a34a',
    text: '#1f2937',
    muted: '#6b7280',
    background: '#ffffff',
    surface: '#f3f4f6',
    border: '#e5e7eb',
  },
  fonts: {
    heading: 'Poppins, Arial, sans-serif',
    body: 'Inter, Arial, sans-serif',
    code: 'JetBrains Mono, monospace',
  },
  fontSize: {
    xs: 8,
    sm: 10,
    base: 11,
    lg: 14,
    xl: 18,
  },
  spacing: {
    xs: 4,
    sm: 8,
    md: 16,
    lg: 24,
    xl: 32,
  },
}

const doc = createDocx({
  theme: brand,
  styles: defineStyles({
    title: {
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      fontSize: '$fontSize.xl',
      marginBottom: '$spacing.lg',
    },
    body: {
      color: '$colors.text',
      fontFamily: '$fonts.body',
      fontSize: '$fontSize.base',
    },
  }),
})
```

## Token Reference Syntax

Reference any token in any style-rule field:

```ts
defineStyles({
  card: {
    color: '$colors.text', // → '#1f2937'
    backgroundColor: '$colors.surface',
    fontFamily: '$fonts.body',
    fontSize: '$fontSize.base',
    margin: '$spacing.md', // shorthand — sets all sides
    border: {
      color: '$colors.border',
      style: 'single',
      width: 1,
    },
  },
  calloutSuccess: {
    color: '$colors.success',
    backgroundColor: '$colors.surface',
    borderColor: '$colors.success',
  },
})
```

## Combining with Presets

Presets define concrete style rules; themes define tokens. Pair them when you want a preset's typography combined with a theme's palette. The modern preset uses primary and accent color tokens; classic and academic headings stay black. Preset font sizes, body fonts, and spacing take precedence over theme defaults.

```ts
import { createDocx, modernPreset, useTheme, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  styles: {
    ...modernPreset.config.styles,
    // Use theme tokens for an extra layer
    highlight: {
      color: '$colors.accent',
      backgroundColor: '$colors.surface',
    },
  },
})
```

## Preview common Word layouts

[Theme Studio](https://docx-kit-theme-studio.ntnyq.dev/?lang=en-US) offers business report, Chinese formal, and academic layouts. Each layout includes representative content and styled tables. The studio explicitly binds the layout's fonts, main type sizes, and heading spacing to the editable theme tokens; the Export theme drawer includes that configuration. Use **Download Word** to save the same file shown in the preview. The preview preserves the exported page margins; installed fonts and the Word layout engine may still differ from browser rendering.

## Next Steps

- [Presets](/guide/presets) — Pre-configured style bundles
- [Styling](/guide/styling) — Full style-system reference
- [Examples: Theme System](/examples/theme-system) — End-to-end themed report
