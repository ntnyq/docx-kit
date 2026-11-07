# 封面（Cover Page）

生成专业标题页或封面，包含标题、副标题、作者、日期、组织，以及可选装饰横线。

## 导入

```ts
import { coverPagePlugin, type CoverPageOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `title` | `string` | _必填_ | 主标题文字 |
| `alignment` | `AlignmentType` | `CENTER` | 文字水平对齐 |
| `author` | `string` | — | 作者或部门名称 |
| `backgroundColor` | `string` | — | RRGGBB 十六进制背景色 |
| `date` | `string` | — | 日期字符串，如 "2026-06-11" |
| `organization` | `string` | — | 组织或公司名称 |
| `showRule` | `boolean` | `true` | 在标题与作者之间显示装饰横线 |
| `subtitle` | `string` | — | 显示在主标题下方的副标题 |

## 示例

### 简单标题页

```ts
import { createDocx, coverPagePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Q3 Operations Report',
    subtitle: 'Data-Driven · Intelligent Decisions',
    author: 'Data Analytics Department',
  })
  .pageBreak()
  .h1('Report Body')
  .p('Content starts here...')
  .save('simple-cover.docx')
```

### 完整封面

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Annual Report 2026',
    subtitle: 'Building the Future Together',
    author: 'Strategic Development Department',
    organization: 'XX Technology Group',
    date: '2026-06-11',
  })
  .pageBreak()
  .h1('Table of Contents')
  .save('full-cover.docx')
```

### 设置背景色

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Confidential',
    subtitle: 'For Internal Use Only',
    author: 'Security Team',
    backgroundColor: 'F0F0F0',
    showRule: false,
  })
  .pageBreak()
  .h1('Contents')
  .save('confidential-cover.docx')
```

### 右对齐封面

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Project Proposal',
    subtitle: 'Phase II Implementation',
    author: 'Engineering Team',
    date: '2026-06-11',
    alignment: AlignmentType.RIGHT,
  })
  .save('right-cover.docx')
```

### 极简封面（仅标题与日期）

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Weekly Status Report',
    subtitle: 'June 11, 2026',
  })
  .save('minimal-cover.docx')
```

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `CoverPageConfig` 和 `CoverPageStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `author`, `date`, `organization`, `subtitle`, `title` | `TextStyleRule` | 文字排版 |
| `authorParagraph`, `dateParagraph`, `organizationParagraph`, `ruleParagraph`, `subtitleParagraph`, `titleParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

```ts
coverPagePlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
