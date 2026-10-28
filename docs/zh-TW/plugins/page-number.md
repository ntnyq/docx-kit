# 頁碼（Page Number）

產生包含頁碼域的段落，適合放在節的頁眉或頁腳中。

## 匯入

```ts
import { pageNumberPlugin, type PageNumberOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `alignment` | `AlignmentType` | `CENTER` | 頁碼水平對齊方式 |
| `fontSize` | `number` | `20` | 字級，單位為半磅（20 = 10pt） |
| `showTotal` | `boolean` | `false` | 顯示 "Page X of Y"，而不只顯示 "X" |

## 範例

### 簡單頁碼

```ts
import { createDocx, pageNumberPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(pageNumberPlugin())
  .plugin('pageNumber', {})
  .save('page-number.docx')
```

### "Page X of Y" 格式

```ts
const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Report')
  .p('This report contains multiple pages.')
  .plugin('pageNumber', {
    showTotal: true,
  })
  .save('page-x-of-y.docx')
```

### 左對齊頁碼

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Document')
  .plugin('pageNumber', {
    alignment: AlignmentType.LEFT,
    fontSize: 18, // 9pt
  })
  .save('left-pn.docx')
```

### 自訂字級

```ts
const doc = createDocx()
  .use(pageNumberPlugin())
  .h1('Large Print Document')
  .plugin('pageNumber', {
    fontSize: 28, // 14pt
    showTotal: true,
  })
  .save('large-pn.docx')
```

### 整合頁眉與頁腳

頁碼外掛適合在節的頁眉或頁腳中使用。以下為概念範例：

```ts
// Note: The plugin output (a Paragraph) can be used in section
// header/footer configuration. The exact API depends on how
// sections with headers/footers are set up.
```
