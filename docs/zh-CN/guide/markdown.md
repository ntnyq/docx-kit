# Markdown

`markdownToNodes()` 将 Markdown 转换为 docx-kit 现有的文档 DSL。
生成的节点与手写节点使用相同的构建器、样式、主题、预设、图片
解析器和插件。转换器基于
[Marked](https://marked.js.org/)，同时支持浏览器与 Node.js。

## 转换与渲染

```ts
import { markdownToNodes } from 'docx-kit/markdown'
import { codeBlockPlugin, createDocx, defineStyles } from 'docx-kit/node'

const markdown = `# Project report

Hello **docx-kit**. Visit [the docs](https://docx-kit.ntnyq.dev).

- [x] Parse Markdown
- [ ] Review the document

~~~ts
const answer = 42
~~~
`

const doc = createDocx({
  styles: defineStyles({
    h1: { fontSize: 24, color: '$colors.primary' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),
  theme: { colors: { primary: '#2563eb' } },
}).use(codeBlockPlugin())

for (const node of markdownToNodes(markdown, {
  classNames: { paragraph: 'body' },
})) {
  doc.add(node)
}

await doc.save('report.docx')
```

在浏览器中，从 `docx-kit` 导入构建器与插件，并通过
`await doc.toBlob()` 导出。转换器不依赖文件系统或 DOM。
使用独立包时，从 `@docxkit/markdown` 导入。

节点也可以作为 schema 内容传入：

```ts
import { markdownToNodes } from 'docx-kit/markdown'
import { codeBlockPlugin, renderDocx } from 'docx-kit/node'

const doc = await renderDocx({
  content: markdownToNodes('# Summary\n\nA **formatted** paragraph.'),
  plugins: [{ plugin: codeBlockPlugin(), type: 'inline' }],
})

await doc.save('summary.docx')
```

## 支持的语法

| Markdown | DSL 输出 |
| --- | --- |
| ATX 与 Setext 标题 | `heading`，1–6 级 |
| 段落与强调 | 包含样式化 `text` 片段的 `paragraph` |
| 加粗、斜体、删除线、行内代码 | 已有文本片段样式 |
| 链接、自动链接、引用链接 | 行内 `hyperlink` |
| 图片与引用图片 | 块级或行内 `image` |
| 无序、有序、嵌套与任务列表 | 列表节点、条目层级与原生行内复选框 |
| GFM 表格 | `table`、列对齐与富内容单元格渲染器 |
| 引用块 | 转换后的块，可使用指定类名或默认左缩进 |
| 围栏与缩进代码块 | 名称为 `codeBlock` 的 `plugin` 节点 |
| 分隔线 | `thematicBreak` |
| 硬换行 | 已有的片段级 `style.docx.break` |

代码块要求调用方注册已有的 `codeBlockPlugin()`。
语言取自围栏信息字符串的第一个单词。语法
高亮沿用该插件对可选 `highlight.js` 的支持。

## 选项与样式表类名

`MarkdownOptions` 接收：

| 选项 | 默认值 | 用途 |
| --- | --- | --- |
| `gfm` | `true` | 表格、任务列表、删除线与扩展自动链接 |
| `breaks` | `false` | 启用 GFM 时，将软换行视为硬换行 |
| `classNames` | 无 | 将语法映射到已有样式表类名 |
| `imageSize` | 无 | 默认图片 `width` 和/或 `height` |
| `transform` | 无 | 替换、调整或忽略每个转换后的块 |

类名目标包括 `heading`、`paragraph`、`blockquote`、`list`、`table`、`code`、
`inlineCode`、`link`、`image`、`hr` 和 `html`，值可以是类名字符串或
数组。请在文档已有的样式表中定义这些类名。未显式
指定标题或段落类名时，仍使用正常的 `h1`–`h6` 与 `p` 回退
规则。行内代码默认使用 Consolas；通过 `inlineCode` 类名
可以由样式表指定字体，也可以使用主题令牌。

## 复用自定义插件

`transform` 接收转换后的节点，以及每个来源块的 `{ type, raw }`。
它在子块转换完成后运行。返回 `[]` 可忽略该块，
也可以返回已有插件节点来自定义渲染：

```ts
import { calloutPlugin, createDocx } from 'docx-kit'
import { markdownToNodes } from 'docx-kit/markdown'

const doc = createDocx().use(calloutPlugin())
const nodes = markdownToNodes('> A note for the reader.', {
  transform(nodes, context) {
    if (context.type === 'blockquote') {
      return [{
        name: 'callout',
        options: {
          content: context.raw.replace(/^> ?/gm, ''),
          type: 'info',
        },
        type: 'plugin',
      }]
    }
    return nodes
  },
})

for (const node of nodes) {
  doc.add(node)
}

const blob = await doc.toBlob()
```

## DSL 边界

- 标题与表头保留可见文本，其现有 DSL
  字段不支持行内格式。
- 表格正文单元格通过列渲染函数保留强调、链接与
  图片。序列化为 JSON 时会保留可读文本，但移除这些
  函数。需要富内容表格时应直接渲染。筛选
  或排序已有行对象会保留富内容。编辑后的
  单元格、复制的行和新增行会将当前值渲染为纯
  文本。
- 图片保留来源与替代文本。URL 或其他自定义来源需要配置 `resolveImage`，
  支持的本地路径可使用 Node.js 入口。
  转换过程不加载图片。链接中的图片会变为替代文本，
  因为超链接的子节点只接受文本片段。
- 原始 HTML 保留为字面文本，不会作为文档标记解析。
- 列表沿用 Word 编号模型，最大嵌套层级为 8，遇到
  嵌套块时拆分列表并保留内容顺序。后续块是
  普通 DSL 块，不是嵌套列表容器。
- 片段链接不会自动生成匹配的标题书签。
