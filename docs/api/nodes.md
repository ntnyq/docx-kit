# DSL Node Types

All node types used in the document content tree. These are the building blocks for both the fluent Builder API and the JSON `renderDocx()` schema.

## `BlockNode<TStyles>`

Union of all top-level content nodes.

```ts
type BlockNode<TStyles> =
  | BookmarkNode<TStyles>
  | CheckboxNode<TStyles>
  | CommentNode<TStyles>
  | FootnoteNode<TStyles>
  | HeadingNode<TStyles>
  | ParagraphNode<TStyles>
  | ImageNode<TStyles>
  | MathNode
  | TableNode<Record<string, unknown>, TStyles>
  | PluginNode<string, unknown, TStyles>
  | HyperlinkNode<TStyles>
  | BulletListNode<TStyles>
  | NumberedListNode<TStyles>
  | RevisionNode<TStyles>
  | ColumnBreakNode
  | PageBreakNode
  | SectionBreakNode
  | TextBoxNode<TStyles>
  | ThematicBreakNode<TStyles>
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
type InlineNode<TStyles> =
  | BookmarkNode<TStyles>
  | CheckboxNode<TStyles>
  | CommentNode<TStyles>
  | FootnoteNode<TStyles>
  | HyperlinkNode<TStyles>
  | ImageNode<TStyles>
  | MathNode
  | RevisionNode<TStyles>
  | TextNode<TStyles>
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

## `TableColumn<TData>`

```ts
interface TableColumn<TData> {
  key: string
  title: string
  align?: 'left' | 'center' | 'right'
  colSpan?: number
  rowSpan?: number
  width?: UnitValue
  cellStyle?: DocxStyleRule | TableCellStyleResolver<TData>
  headerCellStyle?: DocxStyleRule
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

`styleName` and `tableLook` pass native Word table style information through
to OOXML. `layout`, `width`, `alignment`, `borders`, and `floating` expose
fixed/autofit layout, outer/inner borders, and floating placement. Cell styles
can be static or computed from `(value, row, rowIndex, column)`, and a data row
may provide a final `_${key}_style` override.

## `HyperlinkNode<TStyles>`

```ts
interface HyperlinkNode<TStyles> extends BaseNode<TStyles> {
  type: 'hyperlink'
  url?: string       // external target
  anchor?: string    // internal BookmarkNode name
  children: (string | TextNode<TStyles>)[]
}
```

## Semantic and Advanced Content

Bookmarks and internal links:

```ts
doc
  .bookmark('details', ['Details'])
  .internalLink('details', 'Jump to details')
```

Checkboxes are Word content controls:

```ts
doc.checkbox({ checked: true, label: 'Approved', alias: 'approval' })
```

Footnotes and comments register their bodies in the corresponding document
parts while emitting references into the content flow:

```ts
doc
  .footnote([
    'This text is written to word/footnotes.xml.',
  ])
  .comment({
    author: 'Ada Lovelace',
    initials: 'AL',
    date: '2026-07-26T00:00:00Z',
    children: [{ type: 'text', text: 'Annotated text' }],
    comment: ['Review this passage.'],
  })
```

Math nodes build native OMML. Expressions support text, fractions, functions,
integrals, radicals, scripts, and sums:

```ts
doc.math([
  {
    type: 'fraction',
    numerator: [{ type: 'text', text: '1' }],
    denominator: [{ type: 'text', text: '2' }],
  },
])
```

Tracked insertions and deletions carry Word revision metadata. Adding either
node enables tracked revisions automatically:

```ts
doc
  .insertedText({
    author: 'Ada',
    date: '2026-07-26T00:00:00Z',
    revisionId: 1,
    children: ['new text'],
  })
  .deletedText({
    author: 'Ada',
    date: '2026-07-26T00:00:00Z',
    revisionId: 2,
    children: ['old text'],
  })
```

Text boxes and thematic breaks are block nodes:

```ts
doc
  .textBox({
    box: { width: '180pt', height: '60pt', position: 'absolute' },
    text: 'Sidebar',
  })
  .thematicBreak()
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
