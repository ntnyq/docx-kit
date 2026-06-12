# Config Types

Document-level configuration and section/page types.

## `DocxKitConfig<TStyles>`

Top-level configuration passed to `createDocx()`.

```ts
interface DocxKitConfig<TStyles extends StyleSheet = StyleSheet> {
  /** Page dimensions and margins. */
  page?: PageConfig
  /** Named style classes (class → style rule map). */
  styles?: TStyles
  /** Semantic design tokens for theming. */
  theme?: DocxTheme
  /** Default styles applied as base for each element type. */
  defaults?: {
    cell?: DocxStyleRule
    image?: DocxStyleRule
    paragraph?: DocxStyleRule
    table?: DocxStyleRule
    text?: DocxStyleRule
  }
  /** OOXML core properties (File → Info). */
  metadata?: {
    creator?: string
    description?: string
    keywords?: string[]
    lastModifiedBy?: string
    subject?: string
    title?: string
  }
}
```

## `PageConfig`

```ts
interface PageConfig {
  size?: PageSize | { width: UnitValue; height: UnitValue }
  orientation?: 'portrait' | 'landscape'
  margin?:
    | UnitValue
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
}
```

## `PageSize`

```ts
type PageSize = 'A3' | 'A4' | 'Legal' | 'Letter'
```

## `DocxTheme`

Theme tokens that can be referenced in styles via `$category.key` syntax.

```ts
interface DocxTheme {
  id?: string
  name?: string
  colors?: ThemeColors
  fonts?: ThemeFonts
  fontSize?: ThemeFontSize
  spacing?: ThemeSpacing
}
```

### `ThemeColors`

Color palette tokens (name → hex color).

```ts
type ThemeColors = Record<string, string>
```

### `ThemeFonts`

Font family tokens (name → font family).

```ts
type ThemeFonts = Record<string, string>
```

### `ThemeFontSize`

Font size tokens (name → number in pt).

```ts
type ThemeFontSize = Record<string, number>
```

### `ThemeSpacing`

Spacing tokens (name → number in pt).

```ts
type ThemeSpacing = Record<string, number>
```

### Theme Token Syntax

Tokens are referenced in styles using `$category.key`:

```ts
// theme
const theme: DocxTheme = {
  colors: { primary: '#1a56db', accent: '#f59e0b', muted: '#6b7280' },
  fonts: { heading: 'Georgia', body: 'Inter' },
  spacing: { xs: 4, sm: 8, md: 16, lg: 24, xl: 48 },
}

// styles using tokens
const styles = defineStyles({
  title: { color: '$colors.primary', fontFamily: '$fonts.heading' },
  card: { marginBottom: '$spacing.lg' },
})
```

Built-in themes available via `useTheme()`:

```ts
import { useTheme } from 'docx-kit'

const ocean = useTheme('ocean')   // Deep blue / teal color scheme
const warm = useTheme('warm')     // Warm earth-tone color scheme
const minimal = useTheme('minimal') // Clean grayscale + blue accent
```

## `SectionConfig`

Multi-section document configuration.

```ts
interface SectionConfig {
  page?: PageConfig
  header?: HeaderFooterConfig
  footer?: HeaderFooterConfig
}
```

## `HeaderFooterConfig`

```ts
interface HeaderFooterConfig {
  default?: HeaderFooterContent   // All pages
  first?: HeaderFooterContent     // First page only
  even?: HeaderFooterContent      // Even pages only
}
```

## `HeaderFooterContent`

```ts
interface HeaderFooterContent {
  children: (string | BlockNode)[]   // String → auto Paragraph, BlockNode → any DSL node
}
```

Header/footer children support both plain strings and any `BlockNode` (including plugin invocations like `pageNumber`):

```ts
doc.section({
  footer: {
    default: {
      children: [
        'Confidential',
        { type: 'plugin', name: 'pageNumber', options: { showTotal: true } },
      ],
    },
  },
})
```
