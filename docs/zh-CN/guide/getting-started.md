# 快速开始

`docx-kit` 是默认推荐入口，它会统一导出核心 API、内置插件、预设与主题。

## 安装

```bash
pnpm add docx-kit
```

## 第一个文档

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

## 选型建议

- 应用开发直接使用 `docx-kit`
- 插件或底层扩展优先使用 `@docxkit/core`
- 浏览器预览使用 `@docxkit/renderer`
- AI 场景查看 `@docxkit/ai` 与 `@docxkit/mcp`

更多包说明见 [包总览](/zh-CN/ecosystem/packages)。
