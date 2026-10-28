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
