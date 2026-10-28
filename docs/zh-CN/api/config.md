# 配置类型

文档级配置，以及节和页面相关类型。

## `DocxKitConfig<TStyles>`

传入 `createDocx()` 的顶层配置。

```ts
interface DocxKitConfig<TStyles extends StyleSheet = StyleSheet> {
  /**
   * Document-level Word feature switches.
   */
  features?: {
    trackRevisions?: boolean
    updateFields?: boolean
  }
  /**
   * Page dimensions and margins.
   */
  page?: PageConfig
  /**
   * Named style classes (class → style rule map).
   */
  styles?: TStyles
  /**
   * Semantic design tokens for theming.
   */
  theme?: DocxTheme
  /**
   * Default styles applied as base for each element type.
   */
  defaults?: {
    cell?: DocxStyleRule
    image?: DocxStyleRule
    paragraph?: DocxStyleRule
    table?: DocxStyleRule
    text?: DocxStyleRule
  }
  /**
   * OOXML core properties (File → Info).
   */
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
  borders?: PageBorderConfig
  footerDistance?: UnitValue
  gutter?: UnitValue
  headerDistance?: UnitValue
  margin?:
    | UnitValue
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
  pageNumber?: {
    format?: 'decimal' | 'lowerLetter' | 'lowerRoman' | 'upperLetter' | 'upperRoman'
    start?: number
  }
  size?: PageSize | { width: UnitValue; height: UnitValue }
  orientation?: 'portrait' | 'landscape'
}

interface PageBorderConfig {
  top?: BorderRule
  right?: BorderRule
  bottom?: BorderRule
  left?: BorderRule
  display?: 'allPages' | 'firstPage' | 'notFirstPage'
  offsetFrom?: 'page' | 'text'
  zOrder?: 'back' | 'front'
}
```

## `PageSize`

```ts
type PageSize = 'A3' | 'A4' | 'Legal' | 'Letter'
```

## `DocxTheme`

样式中可通过 `$category.key` 语法引用的主题令牌。

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

调色板令牌（名称 → 十六进制颜色）。

```ts
type ThemeColors = Record<string, string>
```

### `ThemeFonts`

字体令牌（名称 → 字体名称）。

```ts
type ThemeFonts = Record<string, string>
```

### `ThemeFontSize`

字号令牌（名称 → 以 pt 为单位的数值）。

```ts
type ThemeFontSize = Record<string, number>
```

### `ThemeSpacing`

间距令牌（名称 → 以 pt 为单位的数值）。

```ts
type ThemeSpacing = Record<string, number>
```

### 主题令牌语法

在样式中使用 `$category.key` 引用令牌：

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

通过 `useTheme()` 可用的内置主题：

```ts
import { useTheme } from 'docx-kit'

const ocean = useTheme('ocean')   // Deep blue / teal color scheme
const warm = useTheme('warm')     // Warm earth-tone color scheme
const minimal = useTheme('minimal') // Clean grayscale + blue accent
```

## `SectionConfig`

多节文档配置。

```ts
interface SectionConfig {
  columns?: {
    count?: number
    spacing?: UnitValue
    separator?: boolean
    equalWidth?: boolean
    columns?: { width: UnitValue; spacing?: UnitValue }[]
  }
  footer?: HeaderFooterConfig
  header?: HeaderFooterConfig
  lineNumbers?: {
    countBy?: number
    distance?: UnitValue
    restart?: 'continuous' | 'newPage' | 'newSection'
    start?: number
  }
  page?: PageConfig
  type?: 'continuous' | 'evenPage' | 'nextColumn' | 'nextPage' | 'oddPage'
}
```

在添加内容前调用 `.section(config)` 会配置第一节，
不会创建空白的前置节：

```ts
const doc = createDocx()
  .section({
    columns: { count: 2, separator: true, spacing: '12pt' },
    lineNumbers: { countBy: 5, restart: 'newPage' },
    page: {
      headerDistance: '8mm',
      footerDistance: '10mm',
      pageNumber: { format: 'upperRoman', start: 1 },
    },
    type: 'continuous',
  })
  .p('First column')
  .columnBreak()
  .p('Second column')
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

页眉和页脚的子内容支持纯字符串和任意 `BlockNode`（包括 `pageNumber` 等插件调用）：

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
