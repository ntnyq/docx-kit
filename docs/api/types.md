# Type Reference

All exported types from docx-kit.

## Utility Types

### `UnitValue`

CSS-like length value — bare numbers or explicit unit strings.

```ts
type UnitValue =
  | `${number}%`
  | `${number}cm`
  | `${number}in`
  | `${number}mm`
  | `${number}pt`
  | `${number}px`
  | number
```

### `MaybePromise<T>`

A value that might be synchronous or wrapped in a Promise.

```ts
type MaybePromise<T> = Promise<T> | T
```

### `LiteralUnion<T, U>`

A literal union that still allows arbitrary values (autocomplete-friendly).

```ts
type LiteralUnion<T extends U, U = string> = T | (U & {})
```

### `HexColor`

A hexadecimal CSS color string.

```ts
type HexColor = `#${string}`
```

### `Dict<T>`

Generic dictionary with string keys.

```ts
type Dict<T = unknown> = Record<string, T>
```

## Style Types

### `DocxStyleRule`

The core CSS-like style descriptor. All properties map to familiar CSS names.

```ts
interface DocxStyleRule {
  // Text
  fontSize?: UnitValue
  fontWeight?: FontWeight
  fontStyle?: 'italic' | 'normal'
  fontFamily?: LiteralUnion<'Arial' | 'Calibri' | 'Times New Roman'>
  color?: string | HexColor
  backgroundColor?: string | HexColor
  underline?: boolean | 'single' | 'double'
  strike?: boolean
  allCaps?: boolean
  letterSpacing?: UnitValue

  // Layout
  textAlign?: TextAlign
  verticalAlign?: VerticalAlign
  lineHeight?: number | UnitValue
  textIndent?: UnitValue

  // Box model
  margin?:
    | UnitValue
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
  marginTop?: UnitValue
  marginBottom?: UnitValue
  marginLeft?: UnitValue
  marginRight?: UnitValue

  // Border
  border?: BorderRule
  borderTop?: BorderRule
  borderBottom?: BorderRule
  borderLeft?: BorderRule
  borderRight?: BorderRule

  // Size
  width?: UnitValue
  height?: UnitValue

  // Escape hatch
  docx?: Record<string, unknown>
}
```

### `FontWeight`

```ts
type FontWeight =
  | 'bold'
  | 'normal'
  | 100
  | 200
  | 300
  | 400
  | 500
  | 600
  | 700
  | 800
  | 900
```

### `TextAlign`

```ts
type TextAlign = 'left' | 'center' | 'right' | 'justify'
```

### `VerticalAlign`

```ts
type VerticalAlign = 'top' | 'middle' | 'bottom'
```

### `BorderRule`

```ts
interface BorderRule {
  color?: string | HexColor
  style?: BorderStyle
  width?: UnitValue
}
```

### `BorderStyle`

```ts
type BorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'none'
```

### `StyleSheet`

A map of class name to style rule.

```ts
type StyleSheet = Record<string, DocxStyleRule>
```

## Document Config Types

### `DocxKitConfig<TStyles>`

Top-level configuration passed to `createDocx()`.

```ts
interface DocxKitConfig<TStyles extends StyleSheet = StyleSheet> {
  page?: PageConfig
  styles?: TStyles
  theme?: DocxTheme
  defaults?: {
    text?: DocxStyleRule
    paragraph?: DocxStyleRule
    table?: DocxStyleRule
    cell?: DocxStyleRule
  }
  metadata?: {
    title?: string
    subject?: string
    creator?: string
    description?: string
    lastModifiedBy?: string
    keywords?: string[]
  }
}
```

### `PageConfig`

```ts
interface PageConfig {
  size?: PageSize | { height: UnitValue; width: UnitValue }
  orientation?: 'portrait' | 'landscape'
  margin?:
    | UnitValue
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
}
```

### `PageSize`

```ts
type PageSize = 'A3' | 'A4' | 'Legal' | 'Letter'
```

### `DocxTheme`

```ts
interface DocxTheme {
  colors?: Record<string, string>
  fontFamily?: Record<string, string>
  fontSize?: Record<string, UnitValue>
  spacing?: Record<string, UnitValue>
}
```

## DSL Node Types

### `BlockNode<TStyles>`

Union of all top-level content nodes.

```ts
type BlockNode<TStyles> =
  | HeadingNode<TStyles>
  | ParagraphNode<TStyles>
  | ImageNode<TStyles>
  | TableNode<Record<string, unknown>, TStyles>
  | PluginNode<string, unknown, TStyles>
  | PageBreakNode
```

### `HeadingNode<TStyles>`

```ts
interface HeadingNode<TStyles> extends BaseNode<TStyles> {
  type: 'heading'
  level: 1 | 2 | 3 | 4 | 5 | 6
  text: string
}
```

### `ParagraphNode<TStyles>`

```ts
interface ParagraphNode<TStyles> extends BaseNode<TStyles> {
  type: 'paragraph'
  text?: string
  children?: InlineNode<TStyles>[]
}
```

### `ImageNode<TStyles>`

```ts
interface ImageNode<TStyles> extends BaseNode<TStyles> {
  type: 'image'
  data: string | ArrayBuffer | Blob | Uint8Array
  width?: UnitValue
  height?: UnitValue
  alt?: string
  imageType?: 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp'
  floating?:
    | boolean
    | {
        wrap?: 'square' | 'tight' | 'topAndBottom'
        x?: UnitValue
        y?: UnitValue
      }
}
```

### `TableNode<TData, TStyles>`

```ts
interface TableNode<TData, TStyles> extends BaseNode<TStyles> {
  type: 'table'
  columns: TableColumn<TData>[]
  data: TData[]
  bordered?: boolean
  striped?: boolean
  header?: boolean
  headerCellStyle?: DocxStyleRule
  cellStyle?: DocxStyleRule
}
```

### `TableColumn<TData>`

```ts
interface TableColumn<TData> {
  key: Extract<keyof TData, string>
  title: string
  align?: 'left' | 'center' | 'right'
  width?: UnitValue
  render?: (
    value: TData[keyof TData],
    row: TData,
    index: number,
  ) => string | InlineNode[]
}
```

### `PageBreakNode`

```ts
interface PageBreakNode {
  type: 'pageBreak'
}
```

### `PluginNode<TName, TOptions, TStyles>`

```ts
interface PluginNode<TName, TOptions, TStyles> extends BaseNode<TStyles> {
  type: 'plugin'
  name: TName
  options: TOptions
}
```

### `BaseNode<TStyles>`

```ts
interface BaseNode<TStyles> {
  className?: string | ClassName<TStyles> | ClassName<TStyles>[]
  id?: string
  style?: DocxStyleRule
}
```

### `InlineNode<TStyles>`

```ts
type InlineNode<TStyles> = TextNode<TStyles> | ImageNode<TStyles>
```

### `TextNode<TStyles>`

```ts
interface TextNode<TStyles> extends BaseNode<TStyles> {
  type: 'text'
  text: string
}
```

## Plugin Types

### `DocxPlugin<TName, TOptions>`

```ts
interface DocxPlugin<TName extends string = string, TOptions = unknown> {
  name: TName
  setup?: () => MaybePromise<void>
  render: (
    options: TOptions,
    context: PluginRenderContext,
  ) => MaybePromise<unknown>
}
```

### `PluginRenderContext`

```ts
interface PluginRenderContext {
  config: DocxKitConfig
  compileNode: (node: BlockNode) => Promise<unknown>
  utils: {
    image: {
      fromBlob: (blob: Blob) => Promise<Uint8Array>
      fromDataUrl: (dataUrl: string) => MaybePromise<Uint8Array>
    }
  }
}
```

### `PluginRegistry`

```ts
type PluginRegistry = Record<string, unknown>
```

## Error Types

### `DocxKitError`

```ts
class DocxKitError extends Error {
  readonly code: string | ErrorCode
  readonly cause: unknown
  constructor(code: string | ErrorCode, message: string, cause?: unknown)
}
```

### `ErrorCode`

```ts
type ErrorCode =
  | 'EXPORT_FAILED'
  | 'IMAGE_INVALID_DATA'
  | 'IMAGE_UNKNOWN_TYPE'
  | 'PLUGIN_NOT_REGISTERED'
  | 'PLUGIN_RENDER_FAILED'
  | 'STYLE_INVALID_UNIT'
  | 'STYLE_UNKNOWN_CLASS'
  | 'TABLE_INVALID_COLUMNS'
  | 'UNKNOWN_NODE_TYPE'
```

## Builder Types

### `DocxBuilder<TStyles, TPlugins>`

```ts
class DocxBuilder<TStyles extends StyleSheet, TPlugins extends PluginRegistry> {
  // Content
  h1(text, options?): this
  h2(text, options?): this
  h3(text, options?): this
  h4(text, options?): this
  h5(text, options?): this
  h6(text, options?): this
  p(text, options?): this
  image(options): this
  table(options): this
  pageBreak(): this
  add(node): this

  // Plugins
  use(plugin): DocxBuilder<TStyles, Record<TName, TOptions> & TPlugins>
  plugin(name, options, style?): this

  // Export
  save(filename): Promise<void>
  toBlob(): Promise<Blob>
  toBuffer(): Promise<Uint8Array>
  toBase64(): Promise<string>
  toDocument(): Promise<Document>
  toJSON(): object
}
```

### `DocxSchema<TStyles>`

```ts
interface DocxSchema<TStyles extends StyleSheet = StyleSheet> {
  content: BlockNode<TStyles>[]
  page?: DocxKitConfig<TStyles>['page']
  styles?: TStyles
}
```

## Platform Types

### Node.js (`docx-kit/node`)

```ts
export function saveDocument(doc: Document, filename: string): Promise<void>
export function dataUrlToUint8Array(dataUrl: string): Promise<Uint8Array>
```

### Browser (`docx-kit/browser`)

```ts
export function dataUrlToUint8Array(dataUrl: string): Promise<Uint8Array>
export function normalizeImageData(
  data: unknown,
): Promise<string | ArrayBuffer | Uint8Array>
```

## Plugin Option Types

### `QRCodePluginOptions`

```ts
interface QRCodePluginOptions {
  text: string
  size?: number // default: 128
  errorCorrectionLevel?: 'L' | 'M' | 'Q' | 'H' // default: 'M'
  margin?: number // default: 1
  caption?: string
}
```

### `EChartsPluginOptions`

```ts
interface EChartsPluginOptions {
  option: EChartsOption
  width?: number // default: 640
  height?: number // default: 360
  renderer?: 'canvas' | 'svg' // default: 'canvas'
  imageType?: 'png' | 'svg' // default: 'png'
  caption?: string
}
```
