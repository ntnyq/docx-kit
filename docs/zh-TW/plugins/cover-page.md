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

## 樣式自訂

工廠的 `{ styles }` 設定實例預設值，`options.styles` 覆寫單次呼叫。外掛套件與 docx-kit 入口匯出 `CoverPageConfig` 和 `CoverPageStyles`。優先順序：**內建/主題預設值 → 初始化 styles → 通用 style → 單次部位 styles**。巢狀欄位合併，陣列整體替換，`undefined` 沿用預設，`0` / `false` 生效；支援單位與 `$colors.primary` 等主題 token。

| 部位 | 樣式型別 | 範圍 |
| --- | --- | --- |
| `author`, `date`, `organization`, `subtitle`, `title` | `TextStyleRule` | 文字排版 |
| `authorParagraph`, `dateParagraph`, `organizationParagraph`, `ruleParagraph`, `subtitleParagraph`, `titleParagraph` | `ParagraphStyleRule` | 段落背景、框線、對齊、間距及分頁 |

```ts
coverPagePlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

單次呼叫在 `options.styles` 中使用相同部位名稱。JSON DSL、AI 工具 schema 與 MCP 驗證共用這些契約；未知部位與放錯部位的屬性會被拒絕。
