# DSL 节点类型

文档内容树中的全部节点类型，同时用于链式 Builder API 和 JSON `renderDocx()` schema。

## `BlockNode<TStyles>`

所有顶层内容节点的联合类型。

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

所有可设置样式的节点都共享这些字段。

```ts
interface BaseNode<TStyles> {
  /**
   * One or more stylesheet class names.
   */
  className?: string | ClassName<TStyles> | ClassName<TStyles>[]
  /**
   * Unique identifier for the node.
   */
  id?: string
  /**
   * Inline style overrides (highest priority).
   */
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

用法：
```ts
{ type: 'heading', level: 1, text: 'Introduction' }
```

## `ParagraphNode<TStyles>`

```ts
interface ParagraphNode<TStyles> extends BaseNode<TStyles> {
  type: 'paragraph'
  /**
   * Simple text (automatically wrapped in a TextNode).
   */
  text?: string
  /**
   * Mixed inline content: text runs + inline images.
   */
  children?: InlineNode<TStyles>[]
}
```

段落子节点支持丰富的行内内容：

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

数据行可通过 `_${key}_colSpan` 和
`_${key}_rowSpan` 字段覆盖单元格合并范围。行级 `_rowSpan` 会应用于
该行的每个单元格。

`styleName` 和 `tableLook` 将 Word 原生表格样式信息传递到
OOXML。`layout`、`width`、`alignment`、`borders` 和 `floating` 提供
固定或自动布局、内外边框与浮动定位。单元格样式
可以是静态值，也可以根据 `(value, row, rowIndex, column)` 计算；数据行
还可以通过 `_${key}_style` 提供最终覆盖。

## `HyperlinkNode<TStyles>`

```ts
interface HyperlinkNode<TStyles> extends BaseNode<TStyles> {
  type: 'hyperlink'
  url?: string       // external target
  anchor?: string    // internal BookmarkNode name
  children: (string | TextNode<TStyles>)[]
}
```

## 语义内容与高级内容

书签与内部链接：

```ts
doc
  .bookmark('details', ['Details'])
  .internalLink('details', 'Jump to details')
```

复选框使用 Word 内容控件：

```ts
doc.checkbox({ checked: true, label: 'Approved', alias: 'approval' })
```

脚注和批注会将正文注册到相应的文档部件，
并在内容流中输出引用：

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

数学节点生成原生 OMML，支持文本、分数、函数、
积分、根式、上下标与求和：

```ts
doc.math([
  {
    type: 'fraction',
    numerator: [{ type: 'text', text: '1' }],
    denominator: [{ type: 'text', text: '2' }],
  },
])
```

修订插入和删除节点带有 Word 修订元数据。添加任意一种
节点都会自动启用修订跟踪：

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

文本框和分隔线属于块级节点：

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

由 `.section()` 创建、用于拆分节的内部节点。

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

列表项可以是纯字符串，也可以是带有样式化行内子节点
和可选嵌套层级的结构化条目。

```ts
interface BulletItem<TStyles> extends BaseNode<TStyles> {
  text?: string
  children?: InlineNode<TStyles>[]
  level?: number // 0–8; overrides the list-level default
}
```

## `ClassName<TStyles>`

从样式表类型中提取有效的类名。

```ts
type ClassName<TStyles extends StyleSheet> = keyof TStyles & string
```
