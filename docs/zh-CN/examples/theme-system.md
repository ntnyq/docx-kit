# 示例：主题系统

展示内置主题系统，包括语义化设计令牌、`$category.key` 语法与三个内置主题（`minimal`、`ocean`、`warm`）。

## 完整代码

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
| `useTheme('ocean')` | 按 ID 加载内置主题 |
| `$colors.primary` 令牌语法 | 从主题解析标题颜色 |
| `$fonts.heading` 令牌语法 | 从主题解析标题字体 |
| `$spacing.md` 令牌语法 | 从主题解析标题外边距 |
| `$fontSize.lg`、`$fontSize.base`、`$fontSize.sm` | 主题字号体系 |
| `$colors.muted`、`$colors.text`、`$colors.info` | 主题调色板中的颜色令牌 |
| `theme` 配置属性 | 将主题传入 `createDocx()` |
| 带令牌的 `defineStyles()` | 样式定义无需绑定特定主题 |

## 可用的内置主题

| 主题 ID | 名称 | 配色 |
|---|---|---|
| `minimal` | Minimal | 灰阶与蓝色强调色 |
| `ocean` | Ocean | 蓝绿色海洋色调 |
| `warm` | Warm | 橙色与琥珀色暖色调 |

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

## 切换主题

只需修改**一行**，即可切换整份文档的外观：

```ts
// Pick one:
const theme = useTheme('minimal')  // Clean grayscale
const theme = useTheme('ocean')    // Blue-green ocean
const theme = useTheme('warm')     // Warm amber/orange
```

同一份 `styles` 对象适用于任意主题，无需修改样式。
