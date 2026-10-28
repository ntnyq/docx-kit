# 样式预设

样式预设是为特定文档类型配置好的 `styles` + `defaults`（以及可选 `theme`）组合，可作为起点并按需覆盖。

## 什么是预设？

预设是一个冻结的 `DocxPreset` 对象，结构如下：

```ts
interface DocxPreset {
  id: string // e.g. 'modern'
  name: string // e.g. 'Modern'
  description: string
  config: DocxKitConfig // styles + defaults + theme
}
```

预设涵盖 `DocxKitConfig` 的 `styles` 与 `defaults` 字段。通常将预设的 `config` 展开传入 `createDocx()`，再叠加自定义配置。

## 内置预设

docx-kit 提供 **3 个内置预设**：

| 预设 | ID | 适用场景 | 字体 |
| ---------------- | ---------- | ---------------------------------------- | --------------- |
| `classicPreset` | `classic` | 中文报告与公文信函 | SimHei / SimSun |
| `modernPreset` | `modern` | 商务报告、专业文档 | Calibri |
| `academicPreset` | `academic` | 学术论文、毕业论文 | Times New Roman |

也可以通过 `usePreset(id)` 按 ID 查找，或通过 `PRESET_LIST` 遍历所有内置预设。

```ts
import {
  usePreset,
  PRESET_LIST,
  classicPreset,
  modernPreset,
  academicPreset,
} from 'docx-kit'

// All three are exported by name:
const modern = modernPreset
const classic = classicPreset
const academic = academicPreset

// Or look up by ID:
const found = usePreset('modern') // returns DocxPreset | undefined

// Or iterate all built-ins:
for (const p of PRESET_LIST) {
  console.log(p.id, p.name)
}
// classic Modern-style  academic
```

## 使用预设

最简单的用法是将预设 `config` 展开传入 `createDocx()`：

```ts
import { createDocx, modernPreset } from 'docx-kit/node'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')
  .p('Revenue grew 15% year-over-year.')
  .p('Key wins in the enterprise segment.')

await doc.save('report.docx')
```

文档会全文使用 Calibri、海军蓝与蓝色标题、通过留白区隔内容、1.35 倍行距，以及垂直外边距为 10pt 的居中图片。

## 合并预设与自定义配置

先展开预设，再覆盖指定字段：

```ts
import { createDocx, modernPreset, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  styles: {
    ...modernPreset.config.styles,
    customClass: { color: '#f00', fontWeight: 'bold' },
  },
  metadata: {
    title: 'My Report',
    creator: 'Alice',
  },
})
```

## 预设详解

### `classicPreset` — 公文样式

适合报告、通知和公文信函的正式中文布局。

```ts
import { createDocx, classicPreset } from 'docx-kit'

const doc = createDocx(classicPreset.config)

doc
  .h1('通知')
  .p('为深入贯彻落实……')
  .h2('一、总体要求')
  .p('以习近平新时代中国特色社会主义思想为指导……')
```

| 属性 | 值 |
| --------- | -------------------------------- |
| `h1` 字体 | 黑体，22pt，加粗，居中 |
| `h2` 字体 | 黑体，16pt，加粗 |
| `h3` 字体 | 楷体，16pt，加粗 |
| 正文字体 | 宋体，14pt，1.5 倍行距 |
| 段落 | 首行缩进 28pt（2 个字符） |
| 图片 | 居中，垂直外边距 8pt |

### `modernPreset` — 商务样式

简洁专业，采用蓝色强调色的商务样式。

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 2026 Report')
  .h2('Executive Summary')
  .p('Revenue exceeded projections by 12%.')
```

| 属性 | 值 |
| --------- | ---------------------------------------- |
| `h1` 字体 | Calibri，24pt，加粗，主题主色 |
| `h2` 字体 | Calibri，16pt，加粗，主题强调色 |
| `h3` 字体 | Calibri，13pt，加粗，主题强调色 |
| 正文字体 | Calibri，11pt，1.35 倍行距 |
| 图片 | 居中，垂直外边距 10pt |

### `academicPreset` — 学术与论文样式

正式学术样式，采用双倍行距与两端对齐。

```ts
import { createDocx, academicPreset } from 'docx-kit'

const doc = createDocx(academicPreset.config)

doc
  .h1('Abstract')
  .p('This paper proposes a novel approach to …')
  .h2('1. Introduction')
  .p('The field of …')
```

| 属性 | 值 |
| --------- | --------------------------------------------------- |
| `h1` 字体 | Times New Roman，16pt，加粗，居中 |
| `h2` 字体 | Times New Roman，14pt，加粗 |
| `h4` 字体 | Times New Roman，12pt，加粗与斜体 |
| 正文字体 | Times New Roman，12pt，**双倍行距**，两端对齐 |
| 段落 | 首行缩进 24pt（2 个字符） |
| 图片 | 居中，垂直外边距 12pt |

## 结合主题

预设和主题互相独立。预设定义_样式规则_（h1 颜色、字号等），主题定义_设计令牌_（调色板、字体栈、尺寸体系），样式通过 `$category.key` 引用这些令牌。

```ts
import { createDocx, modernPreset, useTheme } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  // theme tokens (e.g. $colors.primary) override literal colors when
  // a style rule references them, but presets don't currently
  // reference tokens — so use theme with custom styles for full effect
})
```

详见[主题](/zh-CN/guide/themes)。

## 后续阅读

- [主题](/zh-CN/guide/themes) — 将预设与主题令牌配合使用
- [样式](/zh-CN/guide/styling) — 掌握样式系统
- [示例：现代预设](/zh-CN/examples/preset-modern) — 完整商务报告示例
- [示例：学术预设](/zh-CN/examples/preset-academic) — 完整论文样式示例

## 页面布局与 Word 编辑

所有内置预设均使用 A4。Modern 的上下边距为 22 mm、左右为 24 mm；Classic 的上下边距为 25 mm、左右为 28 mm；Academic 的四周边距为 25.4 mm。标题与后续段落尽量保持同页，并重置首行缩进。正文缩进不会影响插件表格或题注。

原生 Word 文档默认值和标题样式可在编辑导出文件时保留排版。拉丁字体与东亚字体分别由 `fontFamily` 和 `fontFamilyEastAsia` 配置。打开文档的机器需要安装相应字体，否则 Word 会替换字体。
