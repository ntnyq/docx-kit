# Builder API

`DocxBuilder` 类、`createDocx()` 工厂函数、`renderDocx()` JSON 入口与导出方法。

## `createDocx(config?)`

主要入口，返回支持链式构建内容的 `DocxBuilder`。

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx({
  page: { size: 'A4', margin: '2cm 2.5cm' },
  styles: { /* defineStyles(...) */ },
  theme: { /* theme tokens */ },
  defaults: {
    text: { fontFamily: 'Arial', fontSize: 11 },
  },
  metadata: {
    title: 'Annual Report',
    creator: 'Alice',
    customProperties: {
      Department: 'Finance',
    },
  },
})
```

二进制字体数据只需嵌入一次，即可在
文档样式中通过字体名称引用：

```ts
import { readFile } from 'node:fs/promises'
import { createDocx } from 'docx-kit/node'

const doc = createDocx({
  fonts: [
    {
      name: 'Inter',
      data: await readFile('./Inter-Regular.ttf'),
    },
  ],
  defaults: {
    text: { fontFamily: 'Inter' },
  },
})
```

## `DocxBuilder<TStyles, TPlugins>`

`createDocx()` 返回的链式构建器类。

```ts
class DocxBuilder<
  TStyles extends StyleSheet = StyleSheet,
  TPlugins extends PluginRegistry = PluginRegistry,
> {
  // ── Content ──
  h1(text: string, opts?: StyleOptions): this
  h2(text: string, opts?: StyleOptions): this
  h3(text: string, opts?: StyleOptions): this
  h4(text: string, opts?: StyleOptions): this
  h5(text: string, opts?: StyleOptions): this
  h6(text: string, opts?: StyleOptions): this
  p(text: string | InlineNode[], opts?: StyleOptions): this
  bookmark(name: string, children: (string | TextNode)[], opts?: StyleOptions): this
  bulletList(items: BulletItem[], opts?: StyleOptions): this
  checkbox(options?: CheckboxOptions): this
  comment(options: CommentOptions): this
  footnote(content: (string | ParagraphNode)[], opts?: StyleOptions): this
  math(children: MathExpression[]): this
  numberedList(items: BulletItem[], opts?: StyleOptions): this
  hyperlink(url: string, text: string, opts?: StyleOptions): this
  internalLink(anchor: string, text: string, opts?: StyleOptions): this
  insertedText(options: RevisionOptions): this
  deletedText(options: RevisionOptions): this
  image(options: ImageNode): this
  table<T>(options: TableOptions<T>): this
  textBox(options: TextBoxOptions): this
  thematicBreak(opts?: StyleOptions): this
  columnBreak(): this
  pageBreak(): this
  section(config?: SectionConfig): this
  add(node: BlockNode): this

  // ── Plugins ──
  use(plugin: DocxPlugin): DocxBuilder<TStyles, TPlugins & Record<TName, TOptions>>
  plugin(name: TName, options: TOptions, opts?: StyleOptions): this

  // ── Rich Content ──
  span(text: string, style?: DocxStyleRule): TextNode
  inlineImg(options: InlineImageOptions): ImageNode

  // ── Debug ──
  toJSON(): object

  // ── Export ──
  save(filename: string): Promise<void>       // Node.js only
  toBlob(): Promise<Blob>
  toUint8Array(): Promise<Uint8Array>
  toBuffer(): Promise<Uint8Array>              // alias for toUint8Array
  toBase64(): Promise<string>
  toDocument(): Promise<Document>              // raw docx Document
}
```

### 内容方法

| 方法 | 说明 |
|---|---|
| `.h1(text, opts?)` | 一级标题 |
| `.h2(text, opts?)` | 二级标题 |
| `.h3(text, opts?)` | 三级标题 |
| `.h4(text, opts?)` | 四级标题 |
| `.h5(text, opts?)` | 五级标题 |
| `.h6(text, opts?)` | 六级标题 |
| `.p(text, opts?)` | 段落，支持可选的行内子节点 |
| `.bookmark(name, children, opts?)` | 命名书签目标 |
| `.bulletList(items, opts?)` | 无序列表 |
| `.checkbox(options?)` | Word 复选框内容控件 |
| `.comment(options)` | 带 Word 批注的内容范围 |
| `.footnote(content, opts?)` | 脚注引用与正文 |
| `.math(children)` | 结构化 Office Math 公式 |
| `.numberedList(items, opts?)` | 有序列表 |
| `.hyperlink(url, text, opts?)` | 可点击的超链接 |
| `.internalLink(anchor, text, opts?)` | 指向文档内书签的链接 |
| `.insertedText(options)` | 修订模式下的插入内容 |
| `.deletedText(options)` | 修订模式下的删除内容 |
| `.table({ columns, data, ... })` | 数据表格 |
| `.image({ data, width?, height?, ... })` | 嵌入图片 |
| `.textBox(options)` | 可定位的 Word 文本框 |
| `.thematicBreak(opts?)` | 水平分隔线 |
| `.columnBreak()` | 强制切换到本节的下一栏 |
| `.pageBreak()` | 强制分页 |
| `.section(config?)` | 开始新节 |
| `.plugin(name, options, opts?)` | 调用已注册插件 |
| `.use(plugin)` | 注册插件 |
| `.add(node)` | 添加原始 DSL 节点 |

### `span()` 与 `inlineImg()` 辅助函数

用于创建 `.p()` 中的行内内容节点：

```ts
import { span, inlineImg } from 'docx-kit'

// span: creates a TextNode
doc.p([
  span('Hello '),
  span('world', { bold: true, color: '#e11d48' }),
  span('!'),
])

// inlineImg: creates an inline ImageNode
doc.p([
  span('Icon: '),
  inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
  span(' done!'),
])
```

### 表格选项

```ts
interface TableOptions<TData> {
  columns: {
    key: string
    title: string
    align?: 'left' | 'center' | 'right'
    width?: UnitValue
    cellStyle?: DocxStyleRule | TableCellStyleResolver<TData>
    headerCellStyle?: DocxStyleRule
    render?: (value: unknown, row: TData, index: number) => string | InlineNode[]
  }[]
  data: TData[]
  alignment?: 'left' | 'center' | 'right'
  bordered?: boolean
  borders?: TableBordersConfig
  floating?: TableFloatingOptions
  striped?: boolean
  header?: boolean
  layout?: 'autofit' | 'fixed'
  styleName?: string
  tableLook?: TableLookOptions
  visuallyRightToLeft?: boolean
  width?: UnitValue
  headerCellStyle?: DocxStyleRule
  cellStyle?: DocxStyleRule | TableCellStyleResolver<TData>
}
```

通过为两张固定宽度的表格设置不同的水平偏移，
浮动定位可以实现并排表格：

```ts
doc
  .table({
    columns,
    data: leftData,
    layout: 'fixed',
    width: '45%',
    floating: { horizontalAnchor: 'page', x: '0.5in', y: '2in' },
  })
  .table({
    columns,
    data: rightData,
    layout: 'fixed',
    width: '45%',
    floating: { horizontalAnchor: 'page', x: '4.25in', y: '2in' },
  })
```

### 导出方法

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx()

await doc.toBlob()         // → Blob (browser & Node.js)
await doc.toUint8Array()   // → Uint8Array
await doc.toBuffer()       // → Uint8Array (alias)
await doc.toBase64()       // → base64 string
await doc.toStream()       // → Node.js stream (Node.js only)
await doc.save('f.docx')   // → writes to disk (Node.js only)
await doc.toDocument()     // → raw docx Document
doc.toJSON()               // → plain object (debug / serialization)
```

## `renderDocx(schema)`

面向 AI/LLM 的 JSON 文档生成入口。

```ts
import { renderDocx } from 'docx-kit'

const doc = await renderDocx({
  content: [
    { type: 'heading', level: 1, text: 'AI-Generated Report' },
    { type: 'paragraph', text: 'Generated from structured JSON.' },
    {
      type: 'table',
      columns: [
        { key: 'name', title: 'Name' },
        { key: 'value', title: 'Value' },
      ],
      data: [{ name: 'Foo', value: 'Bar' }],
    },
  ],
  styles: { accent: { color: '#2563eb', fontWeight: 'bold' } },
})
const blob = await doc.toBlob()
```

## `DocxSchema<TStyles>`

```ts
interface DocxSchema<TStyles extends StyleSheet = StyleSheet> {
  content: BlockNode<TStyles>[]
  page?: DocxKitConfig<TStyles>['page']
  styles?: TStyles
  plugins?: PluginSource[]
}
```
