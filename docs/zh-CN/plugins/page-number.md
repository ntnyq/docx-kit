# 页码（Page Number）

生成包含页码域的段落，适合放在节的页眉或页脚中。

## 导入

```ts
import { pageNumberPlugin, type PageNumberOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `alignment` | `AlignmentType` | `CENTER` | 页码水平对齐方式 |
| `fontSize` | `number` | `20` | 字号，单位为半磅（20 = 10pt） |
| `showTotal` | `boolean` | `false` | 显示 "Page X of Y"，而不只显示 "X" |

## 示例

### 简单页码

```ts
import { createDocx, pageNumberPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(pageNumberPlugin())
  .plugin('pageNumber', {})
  .save('page-number.docx')
```

### "Page X of Y" 格式

```ts
const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Report')
  .p('This report contains multiple pages.')
  .plugin('pageNumber', {
    showTotal: true,
  })
  .save('page-x-of-y.docx')
```

### 左对齐页码

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Document')
  .plugin('pageNumber', {
    alignment: AlignmentType.LEFT,
    fontSize: 18, // 9pt
  })
  .save('left-pn.docx')
```

### 自定义字号

```ts
const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Large Print Document')
  .plugin('pageNumber', {
    fontSize: 28, // 14pt
    showTotal: true,
  })
  .save('large-pn.docx')
```

### 集成页眉与页脚

页码插件适合在节的页眉或页脚中使用。以下为概念示例：

```ts
// Note: The plugin output (a Paragraph) can be used in section
// header/footer configuration. The exact API depends on how
// sections with headers/footers are set up.
```

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `PageNumberConfig` 和 `PageNumberStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `text` | `TextStyleRule` | 文字排版 |
| `paragraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

```ts
pageNumberPlugin({
  styles: { text: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
