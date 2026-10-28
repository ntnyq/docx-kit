# DSL 節點型別

文件內容樹中的全部節點型別，同時用於鏈式 Builder API 和 JSON `renderDocx()` schema。

## `BlockNode<TStyles>`

所有頂層內容節點的聯合型別。

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

所有可設置樣式的節點都共享這些欄位。

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

段落子節點支援豐富的行內內容：

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

資料行可透過 `_${key}_colSpan` 和
`_${key}_rowSpan` 欄位覆蓋儲存格合併範圍。行級 `_rowSpan` 會應用於
該行的每個儲存格。

`styleName` 和 `tableLook` 將 Word 原生表格樣式資訊傳遞到
OOXML。`layout`、`width`、`alignment`、`borders` 和 `floating` 提供
固定或自動佈局、內外邊框與浮動定位。儲存格樣式
可以是靜態值，也可以根據 `(value, row, rowIndex, column)` 計算；資料行
還可以透過 `_${key}_style` 提供最終覆蓋。

## `HyperlinkNode<TStyles>`

```ts
interface HyperlinkNode<TStyles> extends BaseNode<TStyles> {
  type: 'hyperlink'
  url?: string       // external target
  anchor?: string    // internal BookmarkNode name
  children: (string | TextNode<TStyles>)[]
}
```

## 語義內容與高級內容

書籤與內部連結：

```ts
doc
  .bookmark('details', ['Details'])
  .internalLink('details', 'Jump to details')
```

核取方塊使用 Word 內容控制項：

```ts
doc.checkbox({ checked: true, label: 'Approved', alias: 'approval' })
```

腳註和註解會將正文註冊到相應的文件部件，
並在內容流中輸出引用：

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

數學節點產生原生 OMML，支援文字、分數、函式、
積分、根式、上下標與求和：

```ts
doc.math([
  {
    type: 'fraction',
    numerator: [{ type: 'text', text: '1' }],
    denominator: [{ type: 'text', text: '2' }],
  },
])
```

修訂插入和刪除節點帶有 Word 修訂中繼資料。添加任意一種
節點都會自動啟用修訂跟蹤：

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

文字框和分隔線屬於區塊級節點：

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

由 `.section()` 建立、用於拆分節的內部節點。

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

清單項可以是純字串，也可以是帶有樣式化行內子節點
和可選嵌套層級的結構化條目。

```ts
interface BulletItem<TStyles> extends BaseNode<TStyles> {
  text?: string
  children?: InlineNode<TStyles>[]
  level?: number // 0–8; overrides the list-level default
}
```

## `ClassName<TStyles>`

從樣式表型別中提取有效的類名。

```ts
type ClassName<TStyles extends StyleSheet> = keyof TStyles & string
```
