# 設定型別

文件級設定，以及節和頁面相關型別。

## `DocxKitConfig<TStyles>`

傳入 `createDocx()` 的頂層設定。

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

樣式中可透過 `$category.key` 語法引用的主題令牌。

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

調色板令牌（名稱 → 十六進制顏色）。

```ts
type ThemeColors = Record<string, string>
```

### `ThemeFonts`

字型令牌（名稱 → 字型名稱）。

```ts
type ThemeFonts = Record<string, string>
```

### `ThemeFontSize`

字級令牌（名稱 → 以 pt 為單位的數值）。

```ts
type ThemeFontSize = Record<string, number>
```

### `ThemeSpacing`

間距令牌（名稱 → 以 pt 為單位的數值）。

```ts
type ThemeSpacing = Record<string, number>
```

### 主題令牌語法

在樣式中使用 `$category.key` 引用令牌：

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

透過 `useTheme()` 可用的內建主題：

```ts
import { useTheme } from 'docx-kit'

const ocean = useTheme('ocean')   // Deep blue / teal color scheme
const warm = useTheme('warm')     // Warm earth-tone color scheme
const minimal = useTheme('minimal') // Clean grayscale + blue accent
```

## `SectionConfig`

多節文件設定。

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

在添加內容前呼叫 `.section(config)` 會設定第一節，
不會建立空白的前置節：

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

頁眉和頁腳的子內容支援純字串和任意 `BlockNode`（包括 `pageNumber` 等外掛呼叫）：

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
