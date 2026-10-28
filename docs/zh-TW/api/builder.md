# Builder API

`DocxBuilder` 類、`createDocx()` 工廠函式、`renderDocx()` JSON 入口與匯出方法。

## `createDocx(config?)`

主要入口，回傳支援鏈式建置內容的 `DocxBuilder`。

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

二進制字型資料只需嵌入一次，即可在
文件樣式中透過字型名稱引用：

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

`createDocx()` 回傳的鏈式建構器類。

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

### 內容方法

| 方法 | 說明 |
|---|---|
| `.h1(text, opts?)` | 一級標題 |
| `.h2(text, opts?)` | 二級標題 |
| `.h3(text, opts?)` | 三級標題 |
| `.h4(text, opts?)` | 四級標題 |
| `.h5(text, opts?)` | 五級標題 |
| `.h6(text, opts?)` | 六級標題 |
| `.p(text, opts?)` | 段落，支援可選的行內子節點 |
| `.bookmark(name, children, opts?)` | 命名書籤目標 |
| `.bulletList(items, opts?)` | 無序清單 |
| `.checkbox(options?)` | Word 核取方塊內容控制項 |
| `.comment(options)` | 帶 Word 註解的內容範圍 |
| `.footnote(content, opts?)` | 腳註引用與正文 |
| `.math(children)` | 結構化 Office Math 公式 |
| `.numberedList(items, opts?)` | 有序清單 |
| `.hyperlink(url, text, opts?)` | 可點擊的超連結 |
| `.internalLink(anchor, text, opts?)` | 指向文件內書籤的連結 |
| `.insertedText(options)` | 修訂模式下的插入內容 |
| `.deletedText(options)` | 修訂模式下的刪除內容 |
| `.table({ columns, data, ... })` | 資料表格 |
| `.image({ data, width?, height?, ... })` | 嵌入圖片 |
| `.textBox(options)` | 可定位的 Word 文字框 |
| `.thematicBreak(opts?)` | 水平分隔線 |
| `.columnBreak()` | 強制切換到本節的下一欄 |
| `.pageBreak()` | 強制分頁 |
| `.section(config?)` | 開始新節 |
| `.plugin(name, options, opts?)` | 呼叫已註冊外掛 |
| `.use(plugin)` | 註冊外掛 |
| `.add(node)` | 添加原始 DSL 節點 |

### `span()` 與 `inlineImg()` 輔助函式

用於建立 `.p()` 中的行內內容節點：

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

### 表格選項

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

透過為兩張固定寬度的表格設置不同的水平偏移，
浮動定位可以實現並排表格：

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

### 匯出方法

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

面向 AI/LLM 的 JSON 文件產生入口。

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
