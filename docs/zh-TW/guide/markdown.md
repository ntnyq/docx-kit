# Markdown

`markdownToNodes()` 將 Markdown 轉換為 docx-kit 現有的文件 DSL。
產生的節點與手寫節點使用相同的建構器、樣式、主題、預設、圖片
解析器和外掛。轉換器基於
[Marked](https://marked.js.org/)，同時支援瀏覽器與 Node.js。

## 轉換與呈現

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

在瀏覽器中，從 `docx-kit` 匯入建構器與外掛，並透過
`await doc.toBlob()` 匯出。轉換器不相依檔案系統或 DOM。
使用獨立套件時，從 `@docxkit/markdown` 匯入。

節點也可以作為 schema 內容傳入：

```ts
import { markdownToNodes } from 'docx-kit/markdown'
import { codeBlockPlugin, renderDocx } from 'docx-kit/node'

const doc = await renderDocx({
  content: markdownToNodes('# Summary\n\nA **formatted** paragraph.'),
  plugins: [{ plugin: codeBlockPlugin(), type: 'inline' }],
})

await doc.save('summary.docx')
```

## 支援的語法

| Markdown | DSL 輸出 |
| --- | --- |
| ATX 與 Setext 標題 | `heading`，1–6 級 |
| 段落與強調 | 包含樣式化 `text` 片段的 `paragraph` |
| 加粗、斜體、刪除線、行內程式碼 | 已有文字片段樣式 |
| 連結、自動連結、引用連結 | 行內 `hyperlink` |
| 圖片與引用圖片 | 區塊級或行內 `image` |
| 無序、有序、嵌套與工作清單 | 清單節點、條目層級與原生行內核取方塊 |
| GFM 表格 | `table`、欄對齊與豐富內容儲存格呈現器 |
| 引用區塊 | 轉換後的區塊，可使用指定類名或預設左縮進 |
| 圍欄與縮進程式碼區塊 | 名稱為 `codeBlock` 的 `plugin` 節點 |
| 分隔線 | `thematicBreak` |
| 硬換行 | 已有的片段級 `style.docx.break` |

程式碼區塊要求呼叫方註冊已有的 `codeBlockPlugin()`。
語言取自圍欄資訊字串的第一個單詞。語法
高亮沿用該外掛對可選 `highlight.js` 的支援。

## 選項與樣式表類名

`MarkdownOptions` 接收：

| 選項 | 預設值 | 用途 |
| --- | --- | --- |
| `gfm` | `true` | 表格、工作清單、刪除線與擴充自動連結 |
| `breaks` | `false` | 啟用 GFM 時，將軟換行視為硬換行 |
| `classNames` | 無 | 將語法映射到已有樣式表類名 |
| `imageSize` | 無 | 預設圖片 `width` 和/或 `height` |
| `transform` | 無 | 替換、調整或忽略每個轉換後的區塊 |

類名目標包括 `heading`、`paragraph`、`blockquote`、`list`、`table`、`code`、
`inlineCode`、`link`、`image`、`hr` 和 `html`，值可以是類名字串或
陣列。請在文件已有的樣式表中定義這些類名。未顯式
指定標題或段落類名時，仍使用正常的 `h1`–`h6` 與 `p` 備援
規則。行內程式碼預設使用 Consolas；透過 `inlineCode` 類名
可以由樣式表指定字型，也可以使用主題令牌。

## 重用自訂外掛

`transform` 接收轉換後的節點，以及每個來源區塊的 `{ type, raw }`。
它在子區塊轉換完成後執行。回傳 `[]` 可忽略該區塊，
也可以回傳已有外掛節點來自訂呈現：

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

## DSL 邊界

- 標題與表頭保留可見文字，其現有 DSL
  欄位不支援行內格式。
- 表格正文儲存格透過欄呈現函式保留強調、連結與
  圖片。序列化為 JSON 時會保留可讀文字，但移除這些
  函式。需要豐富內容表格時應直接呈現。篩選
  或排序已有行物件會保留豐富內容。編輯後的
  儲存格、複製的行和新增行會將目前值呈現為純
  文字。
- 圖片保留來源與替代文字。URL 或其他自訂來源需要設定 `resolveImage`，
  支援的本地路徑可使用 Node.js 入口。
  轉換過程不載入圖片。連結中的圖片會變為替代文字，
  因為超連結的子節點只接受文字片段。
- 原始 HTML 保留為字面文字，不會作為文件標記解析。
- 清單沿用 Word 編號模型，最大嵌套層級為 8，遇到
  嵌套區塊時拆分清單並保留內容順序。後續區塊是
  普通 DSL 區塊，不是嵌套清單容器。
- 片段連結不會自動產生匹配的標題書籤。
