# 快速開始

`docx-kit` 是預設推薦入口，會統一重匯出核心 API、內建外掛、預設與主題。

## 安裝

```bash
pnpm add docx-kit
```

## 第一個文件

```ts
import { createDocx, defineStyles } from 'docx-kit'

const doc = createDocx({
  styles: defineStyles({
    title: { fontSize: 24, fontWeight: 'bold', textAlign: 'center' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),
})

doc
  .h1('Hello, docx-kit!', { className: 'title' })
  .p('This is a paragraph.', { className: 'body' })
```

## 選型建議

- 應用開發直接使用 `docx-kit`
- 外掛或底層擴充優先使用 `@docxkit/core`
- 瀏覽器預覽使用 `@docxkit/renderer`
- AI 場景查看 `@docxkit/ai` 與 `@docxkit/mcp`
