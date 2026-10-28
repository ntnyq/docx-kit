# 封面（Cover Page）

產生專業標題頁或封面，包含標題、副標題、作者、日期、組織，以及可選裝飾橫線。

## 匯入

```ts
import { coverPagePlugin, type CoverPageOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `title` | `string` | _必填_ | 主標題文字 |
| `alignment` | `AlignmentType` | `CENTER` | 文字水平對齊 |
| `author` | `string` | — | 作者或部門名稱 |
| `backgroundColor` | `string` | — | RRGGBB 十六進制背景色 |
| `date` | `string` | — | 日期字串，如 "2026-06-11" |
| `organization` | `string` | — | 組織或公司名稱 |
| `showRule` | `boolean` | `true` | 在標題與作者之間顯示裝飾橫線 |
| `subtitle` | `string` | — | 顯示在主標題下方的副標題 |

## 範例

### 簡單標題頁

```ts
import { createDocx, coverPagePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Q3 Operations Report',
    subtitle: 'Data-Driven · Intelligent Decisions',
    author: 'Data Analytics Department',
  })
  .pageBreak()
  .h1('Report Body')
  .p('Content starts here...')
  .save('simple-cover.docx')
```

### 完整封面

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Annual Report 2026',
    subtitle: 'Building the Future Together',
    author: 'Strategic Development Department',
    organization: 'XX Technology Group',
    date: '2026-06-11',
  })
  .pageBreak()
  .h1('Table of Contents')
  .save('full-cover.docx')
```

### 設置背景色

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Confidential',
    subtitle: 'For Internal Use Only',
    author: 'Security Team',
    backgroundColor: 'F0F0F0',
    showRule: false,
  })
  .pageBreak()
  .h1('Contents')
  .save('confidential-cover.docx')
```

### 右對齊封面

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Project Proposal',
    subtitle: 'Phase II Implementation',
    author: 'Engineering Team',
    date: '2026-06-11',
    alignment: AlignmentType.RIGHT,
  })
  .save('right-cover.docx')
```

### 極簡封面（僅標題與日期）

```ts
const doc = createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Weekly Status Report',
    subtitle: 'June 11, 2026',
  })
  .save('minimal-cover.docx')
```
