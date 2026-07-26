# DSL Node Types

All node types used in the document content tree. These are the building blocks for both the fluent Builder API and the JSON `renderDocx()` schema.

## `BlockNode<TStyles>`

Union of all top-level content nodes.

```ts
type BlockNode<TStyles> =
  | HeadingNode<TStyles>
  | ParagraphNode<TStyles>
  | ImageNode<TStyles>
  | TableNode<Record<string, unknown>, TStyles>
  | PluginNode<string, unknown, TStyles>
  | HyperlinkNode<TStyles>
  | BulletListNode<TStyles>
  | NumberedListNode<TStyles>
  | ColumnBreakNode
  | PageBreakNode
  | SectionBreakNode
```

## `BaseNode<TStyles>`

All styled nodes share these common fields.

```ts
interface BaseNode<TStyles> {
  /** One or more stylesheet class names. */
  className?: string | ClassName<TStyles> | ClassName<TStyles>[]
  /** Unique identifier for the node. */
  id?: string
  /** Inline style overrides (highest priority). */
  style?: DocxStyleRule
}
```

## `HeadingNode<TStyles>`

```ts
interface HeadingNode<TStyles> extends BaseNode<TStyles> {
  type: 'heading'
  level: 1 | 2 | 3 | 4 | 5 | 6
  text: string
}
```

Usage:
```ts
{ type: 'heading', level: 1, text: 'Introduction' }
```

## `ParagraphNode<TStyles>`

```ts
interface ParagraphNode<TStyles> extends BaseNode<TStyles> {
  type: 'paragraph'
  /** Simple text (automatically wrapped in a TextNode). */
  text?: string
  /** Mixed inline content: text runs + inline images. */
  children?: InlineNode<TStyles>[]
}
```

Paragraph children support rich inline content:

```ts
// Simple text
{ type: 'paragraph', text: 'Hello world' }

// Mixed inline content
{
  type: 'paragraph',
  children: [
    { type: 'text', text: 'Hello ' },
    { type: 'text', text: 'world', style: { bold: true, color: '#f00' } },
  ],
}

// Inline image
{
  type: 'paragraph',
  children: [
    { type: 'text', text: 'Icon: ' },
    { type: 'image', data: iconUrl, width: 16, height: 16 },
    { type: 'text', text: ' done!' },
  ],
}
```

## `InlineNode<TStyles>`

```ts
type InlineNode<TStyles> = TextNode<TStyles> | ImageNode<TStyles>
```

## `TextNode<TStyles>`

```ts
interface TextNode<TStyles> extends BaseNode<TStyles> {
  type: 'text'
  text: string
}
```

## `ImageNode<TStyles>`

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

## `TableNode<TData, TStyles>`

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

## `TableColumn<TData>`

```ts
interface TableColumn<TData> {
  key: string
  title: string
  align?: 'left' | 'center' | 'right'
  colSpan?: number
  rowSpan?: number
  width?: UnitValue
  render?: (
    value: unknown,
    row: TData,
    index: number,
  ) => string | InlineNode[]
}
```

Data rows can override spans with `_${key}_colSpan` and
`_${key}_rowSpan` fields. A row-level `_rowSpan` applies to every cell in
that row.

## `HyperlinkNode<TStyles>`

```ts
interface HyperlinkNode<TStyles> extends BaseNode<TStyles> {
  type: 'hyperlink'
  url: string
  children: (string | TextNode<TStyles>)[]
}
```

## `BulletListNode<TStyles>`

```ts
interface BulletListNode<TStyles> extends BaseNode<TStyles> {
  type: 'bulletList'
  items: BulletItem<TStyles>[]
  bullet?: string           // Custom bullet character
  level?: number            // Nesting level
}
```

## `NumberedListNode<TStyles>`

```ts
interface NumberedListNode<TStyles> extends BaseNode<TStyles> {
  type: 'numberedList'
  items: BulletItem<TStyles>[]
  numberingFormat?: 'decimal' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman'
  start?: number            // Starting number
  level?: number            // Nesting level
}
```

## `PageBreakNode`

```ts
interface PageBreakNode {
  type: 'pageBreak'
}
```

## `ColumnBreakNode`

```ts
interface ColumnBreakNode {
  type: 'columnBreak'
}
```

## `SectionBreakNode`

Internal node created by `.section()` to split sections.

```ts
interface SectionBreakNode {
  type: 'sectionBreak'
  config?: SectionConfig
}
```

## `PluginNode<TName, TOptions, TStyles>`

```ts
interface PluginNode<TName, TOptions, TStyles> extends BaseNode<TStyles> {
  type: 'plugin'
  name: TName
  options: TOptions
}
```

## `BulletItem<TStyles>`

A list item can be a plain string or a structured item with styled inline
children and an optional per-item nesting level.

```ts
interface BulletItem<TStyles> extends BaseNode<TStyles> {
  text?: string
  children?: InlineNode<TStyles>[]
  level?: number // 0–8; overrides the list-level default
}
```

## `ClassName<TStyles>`

Extracts valid class names from a stylesheet type.

```ts
type ClassName<TStyles extends StyleSheet> = keyof TStyles & string
```
