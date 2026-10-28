# CSS 风格样式

docx-kit 使用 CSS 风格属性名，让网页开发者能直观地设置文档样式。

## 基础样式

样式规则使用熟悉的 CSS 属性名，内部映射到 Word OOXML：

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

### 全部样式属性

| 属性 | 类型 | 说明 |
|---|---|---|
| `fontSize` | `UnitValue` | 字号，纯数字单位为 pt |
| `fontWeight` | `FontWeight` | 字重：`'bold'`、`'normal'` 或 `100`–`900` |
| `fontStyle` | `'italic' \| 'normal'` | 斜体开关 |
| `fontFamily` | `LiteralUnion<'Arial' \| 'Calibri' \| 'Times New Roman'>` | 字体；DOCX 使用 CSS 回退字体栈中的第一个字体 |
| `color` | `string \| HexColor` | 文字或前景颜色 |
| `backgroundColor` | `string \| HexColor` | 背景底纹 |
| `highlight` | `HighlightColor` | 文字高亮（`'yellow'`、`'green'`、`'red'` 等） |
| `textAlign` | `'left' \| 'center' \| 'right' \| 'justify'` | 水平对齐 |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | 垂直对齐（表格单元格） |
| `lineHeight` | `number \| UnitValue` | 行高 |
| `letterSpacing` | `UnitValue` | 字符间距 |
| `textIndent` | `UnitValue` | 首行缩进 |
| `underline` | `boolean \| 'single' \| 'double' \| 'dash' \| 'dotDash'` | 下划线 |
| `strike` | `boolean` | 删除线 |
| `allCaps` | `boolean` | 强制大写 |
| `margin` | `UnitValue \| 简写字符串` | CSS 风格外边距简写 |
| `marginTop/Bottom/Left/Right` | `UnitValue` | 各方向外边距 |
| `border` | `BorderRule` | 四边边框简写 |
| `borderTop/Bottom/Left/Right` | `BorderRule` | 各方向边框 |
| `width` | `UnitValue` | 元素宽度 |
| `height` | `UnitValue` | 元素高度 |
| `docx` | `Record<string, unknown>` | 直接透传给 `docx` 构造函数 |

## 单位值

长度可以使用纯数字或带显式单位的字符串：

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

### 单位约定

| 场景 | 纯数字的单位 |
|---|---|
| `fontSize` | pt |
| `margin*` / `padding*` | pt |
| `textIndent` | pt |
| `letterSpacing` | pt |
| 图片 `width` / `height` | px |

## 命名样式表（defineStyles）

通过 `defineStyles()` 定义可复用样式：

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

通过节点的 `className` 使用：

```ts
doc
  .h1('Title', { className: 'title' })
  .p('Body text', { className: 'body' })
  .p('Disclaimer', { className: 'small' })
  .p('Important!', { className: 'red' })
```

多个类名（空格分隔或数组）：

```ts
doc.p('Important note', {
  className: 'body red',  // space-separated
})
```

## 样式优先级（层叠）

样式按以下优先级合并，高优先级覆盖低优先级：

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

## 边框

按方向设置边框样式、宽度和颜色：

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

可用边框样式：`'single'`、`'double'`、`'dashed'`、`'dotted'`、`'none'`。

## 外边距简写

与 CSS 类似，支持 1、2 或 4 个值的外边距简写：

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

## 底层配置（docx）

对于尚未映射为 CSS 风格属性的 OOXML 选项，可以使用 `docx`：

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

## 各元素类型的默认值

在 `createDocx()` 配置中为每种元素设置默认样式：

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

## 主题令牌

通过 `theme` 选项定义可复用设计令牌，再用 `$category.key` 语法在样式中引用：

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

### 内置主题

docx-kit 提供 3 个内置主题：

```ts
import { useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),   // 'minimal' | 'ocean' | 'warm'
})
```

| 主题 | 配色 | 适用场景 |
|---|---|---|
| `minimal` | 简洁灰阶与蓝色强调色 | 通用文档 |
| `ocean` | 深蓝与青绿色配色 | 商务报告 |
| `warm` | 暖色大地色系 | 创意文档、演示材料 |

## 样式预设

为便于快速开始，docx-kit 提供 **3 个预配置样式预设**。详见[样式预设](/zh-CN/guide/presets)。

使用 `modernPreset` 的简短示例：

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')           // Navy 26pt with blue underline (Calibri)
  .p('Body text here.')      // Calibri 11pt, 1.5× line height
```

## 主题

主题是颜色、字体与尺寸体系等设计令牌，在样式中通过 `$category.key` 引用。docx-kit 提供 **3 个内置主题**（`minimal`、`ocean`、`warm`），完整令牌参考见[主题](/zh-CN/guide/themes)。

使用 Ocean 主题的简短示例：

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

## 相关文档

- [样式预设](/zh-CN/guide/presets) — 预配置样式组合
- [主题](/zh-CN/guide/themes) — 设计令牌
- [Builder API](/zh-CN/guide/builder-api) — 完整构建器参考
- [示例：主题系统](/zh-CN/examples/theme-system) — 完整主题报告

## 样式继承（extends）

样式可以继承其他样式，实现类似 CSS 的继承：

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

继承在文档编译阶段解析，并检测循环引用。

完整演示见[样式继承示例](/zh-CN/examples/style-inheritance)。

## 样式预设

使用预配置样式预设快速开始：

```ts
import { createDocx, usePreset } from 'docx-kit'

// 3 built-in presets available:
const doc = createDocx(usePreset('modern').config)
// or 'classic' (gov-doc style), 'academic' (thesis style)
```

每个预设都为对应场景提供完整样式表、页面配置与元素默认值。也可以将预设与自定义配置合并：

```ts
const doc = createDocx({
  ...usePreset('modern').config,
  styles: {
    ...usePreset('modern').config.styles,
    customClass: { color: '#f00' },
  },
})
```
