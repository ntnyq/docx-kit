# Builder API

The `DocxBuilder` class, `createDocx()` factory, `renderDocx()` JSON entry point, and export methods.

## `createDocx(config?)`

The primary entry point. Returns a fluent `DocxBuilder` for chaining content.

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
  },
})
```

## `DocxBuilder<TStyles, TPlugins>`

The fluent builder class returned by `createDocx()`.

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
  bulletList(items: BulletItem[], opts?: StyleOptions): this
  numberedList(items: BulletItem[], opts?: StyleOptions): this
  hyperlink(url: string, text: string, opts?: StyleOptions): this
  image(options: ImageNode): this
  table<T>(options: TableOptions<T>): this
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

### Content Methods

| Method | Description |
|---|---|
| `.h1(text, opts?)` | Level-1 heading |
| `.h2(text, opts?)` | Level-2 heading |
| `.h3(text, opts?)` | Level-3 heading |
| `.h4(text, opts?)` | Level-4 heading |
| `.h5(text, opts?)` | Level-5 heading |
| `.h6(text, opts?)` | Level-6 heading |
| `.p(text, opts?)` | Paragraph with optional inline children |
| `.bulletList(items, opts?)` | Unordered list |
| `.numberedList(items, opts?)` | Ordered list |
| `.hyperlink(url, text, opts?)` | Clickable hyperlink |
| `.table({ columns, data, ... })` | Data table |
| `.image({ data, width?, height?, ... })` | Embedded image |
| `.columnBreak()` | Forced break to the next section column |
| `.pageBreak()` | Forced page break |
| `.section(config?)` | Start a new section |
| `.plugin(name, options, opts?)` | Invoke registered plugin |
| `.use(plugin)` | Register a plugin |
| `.add(node)` | Add raw DSL node |

### `span()` and `inlineImg()` Helpers

These create inline content nodes for use inside `.p()`:

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

### Table Options

```ts
interface TableOptions<TData> {
  columns: {
    key: string
    title: string
    align?: 'left' | 'center' | 'right'
    width?: UnitValue
    render?: (value: unknown, row: TData, index: number) => string | InlineNode[]
  }[]
  data: TData[]
  bordered?: boolean
  striped?: boolean
  header?: boolean
  headerCellStyle?: DocxStyleRule
  cellStyle?: DocxStyleRule
}
```

### Export Methods

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx()

await doc.toBlob()         // → Blob (browser & Node.js)
await doc.toUint8Array()   // → Uint8Array
await doc.toBuffer()       // → Uint8Array (alias)
await doc.toBase64()       // → base64 string
await doc.save('f.docx')   // → writes to disk (Node.js only)
await doc.toDocument()     // → raw docx Document
doc.toJSON()               // → plain object (debug / serialization)
```

## `renderDocx(schema)`

AI/LLM-friendly JSON entry point for document generation.

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
