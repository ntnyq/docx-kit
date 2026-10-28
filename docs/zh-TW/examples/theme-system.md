# 範例：主題系統

展示內建主題系統，包括語義化設計令牌、`$category.key` 語法與三個內建主題（`minimal`、`ocean`、`warm`）。

## 完整程式碼

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

## 本例展示的能力

| 功能 | 使用位置 |
|---|---|
| `useTheme('ocean')` | 按 ID 載入內建主題 |
| `$colors.primary` 令牌語法 | 從主題解析標題顏色 |
| `$fonts.heading` 令牌語法 | 從主題解析標題字型 |
| `$spacing.md` 令牌語法 | 從主題解析標題外邊距 |
| `$fontSize.lg`、`$fontSize.base`、`$fontSize.sm` | 主題字級體系 |
| `$colors.muted`、`$colors.text`、`$colors.info` | 主題調色板中的顏色令牌 |
| `theme` 設定屬性 | 將主題傳入 `createDocx()` |
| 帶令牌的 `defineStyles()` | 樣式定義無需綁定特定主題 |

## 可用的內建主題

| 主題 ID | 名稱 | 配色 |
|---|---|---|
| `minimal` | Minimal | 灰階與藍色強調色 |
| `ocean` | Ocean | 藍綠色海洋色調 |
| `warm` | Warm | 橙色與琥珀色暖色調 |

## 令牌解析原理

```
defineStyles({
  title: { color: '$colors.primary' }
})

     ↓ compile time

useTheme('ocean') → theme.colors.primary = '#0ea5e9'

     ↓ resolveThemeTokens()

{ color: '#0ea5e9' }
```

## 切換主題

只需修改**一行**，即可切換整份文件的外觀：

```ts
// Pick one:
const theme = useTheme('minimal')  // Clean grayscale
const theme = useTheme('ocean')    // Blue-green ocean
const theme = useTheme('warm')     // Warm amber/orange
```

同一份 `styles` 物件適用於任意主題，無需修改樣式。
