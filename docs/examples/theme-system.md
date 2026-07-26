# Example: Theme System

Showcase the built-in theme system with semantic design tokens, `$category.key` token syntax, and the three built-in themes (`minimal`, `ocean`, `warm`).

## Full Code

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit/node'

// 1. Load a built-in theme
const theme = useTheme('ocean')

// 2. Define styles that reference theme tokens via $category.key syntax
const styles = defineStyles({
  // Token references are resolved at compile time
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '$colors.primary',       // → theme.colors.primary
    fontFamily: '$fonts.heading',    // → theme.fonts.heading
    textAlign: 'center',
    marginBottom: '$spacing.md',     // → theme.spacing.md
  },
  subtitle: {
    fontSize: '$fontSize.lg',
    color: '$colors.muted',
    textAlign: 'center',
    marginBottom: '$spacing.xl',
  },
  body: {
    fontSize: '$fontSize.base',
    color: '$colors.text',
    fontFamily: '$fonts.body',
    lineHeight: 1.6,
  },
  callout: {
    backgroundColor: '$colors.info',
    color: '#ffffff',
    padding: '$spacing.sm',
    borderRadius: 4,
  },
  footer: {
    fontSize: '$fontSize.sm',
    color: '$colors.muted',
    textAlign: 'center',
    marginTop: '$spacing.xl',
  },
})

// 3. Build with theme
const doc = createDocx({
  styles,
  theme,
  page: { size: 'A4', margin: '20mm 25mm' },
  metadata: {
    title: 'Theme System Demo',
    creator: 'docx-kit',
  },
})

doc
  .h1('Theme System Demo', { className: 'title' })
  .p('This document uses the ocean theme with semantic design tokens.', {
    className: 'subtitle',
  })

  .p(
    'Theme tokens are referenced in styles using the $category.key syntax. ' +
    'At compile time, $colors.primary is replaced with the actual color value ' +
    'from the theme (e.g. #2563eb for the ocean theme).',
    { className: 'body' },
  )

  .p(
    '✨ Callout: You can switch themes by changing a single line: ' +
    "useTheme('warm') or useTheme('minimal').",
    { className: 'callout' },
  )

  .p('Footer text uses $fontSize.sm and $colors.muted from the theme.', {
    className: 'footer',
  })

  .save('theme-system-demo.docx')
```

## What This Demonstrates

| Feature | Used In |
|---|---|
| `useTheme('ocean')` | Load built-in theme by ID |
| `$colors.primary` token syntax | Title color resolved from theme |
| `$fonts.heading` token syntax | Title font family resolved from theme |
| `$spacing.md` token syntax | Title margin resolved from theme |
| `$fontSize.lg` / `$fontSize.base` / `$fontSize.sm` | Font sizes from theme scale |
| `$colors.muted` / `$colors.text` / `$colors.info` | Color tokens from theme palette |
| `theme` config property | Pass theme to `createDocx()` |
| `defineStyles()` with tokens | Style definitions stay theme-agnostic |

## Available Built-in Themes

| Theme ID | Name | Palette |
|---|---|---|
| `minimal` | Minimal | Grayscale + blue accent |
| `ocean` | Ocean | Blue-green ocean tones |
| `warm` | Warm | Warm orange/amber tones |

## How Token Resolution Works

```
defineStyles({
  title: { color: '$colors.primary' }
})

     ↓ compile time

useTheme('ocean') → theme.colors.primary = '#0ea5e9'

     ↓ resolveThemeTokens()

{ color: '#0ea5e9' }
```

## Switching Themes

To switch the entire document look, change **one line**:

```ts
// Pick one:
const theme = useTheme('minimal')  // Clean grayscale
const theme = useTheme('ocean')    // Blue-green ocean
const theme = useTheme('warm')     // Warm amber/orange
```

The same `styles` object works with any theme — no style changes needed.
