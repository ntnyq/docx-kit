# 主题

主题是**语义化设计令牌**，包含颜色、字体、字号与间距体系，可通过 `$category.key` 在样式规则中引用，将_样式定义_与_具体颜色和字体_解耦。主题字体还会提供 Word 原生默认值，让未设置样式的文字和插件自动继承文档字体。

## 什么是主题？

`DocxTheme` 是设计令牌的集合：

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

**为什么使用主题？**

- 修改一个字段即可切换整份文档的外观
- 让多份文档保持配色一致
- 按语义复用 `primary`、`accent`、`success` 等令牌
- 与[预设](/zh-CN/guide/presets)配合，得到完整初始配置

## 内置主题

docx-kit 提供 **3 个内置主题**：

| 主题 | ID | 配色 | 适用场景 | 字体 |
| -------------- | --------- | ----------------------------------- | ---------------------- | ----------------- |
| `minimalTheme` | `minimal` | 简洁灰阶与蓝色强调色 | 通用商务文档 | Calibri |
| `oceanTheme` | `ocean` | 海军蓝、青绿色与浅灰底色 | 报告、文章 | Cambria / Calibri |
| `warmTheme` | `warm` | 暖灰与棕色 | 信函、随笔 | Georgia |

辅助函数：

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

## 使用主题

将主题传入 `createDocx()`：

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

样式规则中以 `$` 开头的字符串会作为令牌引用，在文档编译阶段解析。

## 令牌分类

### `$colors.*` — 调色板

每个内置主题提供 12 个语义化颜色令牌：

| 令牌 | 用途 | `minimal` | `ocean` | `warm` |
| ------------ | ---------------------- | --------- | --------- | --------- |
| `primary` | 标题、主要文字 | `#111827` | `#0f172a` | `#44403c` |
| `secondary` | 次要文字 | `#374151` | `#1e293b` | `#57534e` |
| `accent` | 强调色与链接颜色 | `#2563eb` | `#0d9488` | `#8b5e34` |
| `success` | 成功状态 | `#2e7d32` | `#14765c` | `#15803d` |
| `warning` | 警告状态 | `#9a6700` | `#946200` | `#a34c13` |
| `danger` | 错误状态 | `#b91c1c` | `#e11d48` | `#b91c1c` |
| `info` | 信息状态 | `#245a96` | `#24758a` | `#6b6578` |
| `muted` | 弱化文字 | `#6b7280` | `#64748b` | `#78716c` |
| `background` | 页面背景 | `#ffffff` | `#ffffff` | `#ffffff` |
| `surface` | 卡片或面板底色 | `#f9fafb` | `#f0f5f7` | `#f6f3ee` |
| `text` | 正文 | `#1f2937` | `#1e293b` | `#292524` |
| `border` | 边框颜色 | `#e5e7eb` | `#cbd9df` | `#ddd6ca` |

### `$fonts.*` — 字体栈

每个内置主题提供正文、标题和代码字体，以及分别用于中日韩文字的 `bodyEastAsia` 和 `headingEastAsia` 字体。Word 对每种文字系统仅接受一种字体，CSS 回退字体栈会取第一个字体。

| 令牌 | `minimal` | `ocean` | `warm` |
| --------- | ---------- | ---------- | ---------- |
| `heading` | `Calibri`  | `Cambria`  | `Georgia`  |
| `body`    | `Calibri`  | `Calibri`  | `Georgia`  |
| `code`    | `Consolas` | `Consolas` | `Consolas` |

Minimal 和 Ocean 的 `bodyEastAsia`、`headingEastAsia` 使用 Microsoft YaHei。
Warm 的正文使用 SimSun，标题使用 SimHei。这些是 Word 字体名称，
打开文档的应用会在字体缺失时选择替代字体。

### `$fontSize.*` — 字号体系

每个内置主题提供 5 级字号，单位为 pt：

| 令牌 | `minimal` | `ocean` | `warm` |
| ------ | --------- | ------- | ------ |
| `xs`   | 8         | 8       | 8      |
| `sm`   | 9         | 10      | 10     |
| `base` | 11        | 12      | 12     |
| `lg`   | 14        | 16      | 15     |
| `xl`   | 24        | 26      | 24     |

### `$spacing.*` — 间距体系

单位为 pt，可用于 `margin*`、`padding*`、`borderWidth` 等：

| 令牌 | `minimal` | `ocean` | `warm` |
| ----- | --------- | ------- | ------ |
| `xs`  | 4         | 5       | 4      |
| `sm`  | 8         | 10      | 9      |
| `md`  | 16        | 18      | 16     |
| `lg`  | 24        | 28      | 26     |
| `xl`  | 32        | 36      | 34     |

## 自定义主题

将自定义 `DocxTheme` 对象传入 `createDocx()`：

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

## 令牌引用语法

可以在样式规则的任意字段中引用令牌：

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

## 结合预设

预设定义具体样式规则，主题定义令牌。可以将预设排版与主题配色组合使用。Modern 预设使用主色与强调色令牌；Classic 和 Academic 的标题保持黑色。预设字号、正文字体和间距优先于主题默认值。

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

## 预览常见 Word 布局

[主题工作室](https://docx-kit-theme-studio.ntnyq.dev/?lang=zh-CN)提供商务报告、中文公文与学术布局，每种布局都包含代表性内容和带样式表格。工作室会显式将布局字体、主要字号与标题间距绑定到可编辑主题令牌，导出主题侧栏包含相应配置。使用**下载 Word**保存与预览相同的文件。预览保留导出页边距；已安装字体与 Word 布局引擎仍可能使最终显示与浏览器有所不同。

## 后续阅读

- [预设](/zh-CN/guide/presets) — 预配置样式组合
- [样式](/zh-CN/guide/styling) — 完整样式系统参考
- [示例：主题系统](/zh-CN/examples/theme-system) — 完整主题报告
