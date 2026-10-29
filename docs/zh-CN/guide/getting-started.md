# 快速开始

## 安装

```bash
pnpm add docx-kit docx
```

请安装 9.x 系列中不低于 9.15.0 的 `docx` 作为共享对等依赖。构建器、内置插件和自定义插件需要使用同一份 `docx` 实例。

### 可选对等依赖

部分内置插件需要额外的包：

```bash
# For QRCode plugin
pnpm add qrcode

# For ECharts plugin (browser only)
pnpm add echarts
```

## 快速开始

用 3 行代码创建第一个 .docx 文档：

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx()

await doc
  .h1('Hello, World')
  .p('This is my first docx-kit document.')
  .save('hello.docx')
```

## 核心概念

docx-kit 围绕三个核心概念构建：

| 概念 | 说明 |
| ----------- | --------------------------------------------------------------------------- |
| **样式** | 将 CSS 风格规则（fontSize、color、margin 等）组织为样式表 |
| **节点** | 内容构建块：标题、段落、表格、图片和插件 |
| **构建器** | 通过链式 API 组合节点并导出为 .docx |

### 1. 定义样式（CSS 风格）

```ts
import { defineStyles } from 'docx-kit'

const styles = defineStyles({
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#1a1a2e',
    textAlign: 'center',
  },
  subtitle: {
    fontSize: 16,
    color: '#666',
    textAlign: 'center',
  },
  body: {
    fontSize: 12,
    lineHeight: 1.5,
  },
  highlight: {
    backgroundColor: '#fff3cd',
    color: '#856404',
  },
})
```

### 2. 构建内容（链式 API）

```ts
const doc = createDocx({ styles })

await doc
  .h1('Quarterly Report', { className: 'title' })
  .p('Q1 2026 Financial Summary', { className: 'subtitle' })
  .p('Revenue grew 15% year-over-year...', { className: 'body' })
  .p('⚠️ Note: exchange rate fluctuations', { className: 'highlight' })
  .save('report.docx')
```

### 3. 导出

```ts
// Save to disk (Node.js)
await doc.save('output.docx')

// Get as Blob (browser)
const blob = await doc.toBlob()

// Get as Uint8Array (cross-platform)
const bytes = await doc.toBuffer()

// Get as base64 string
const b64 = await doc.toBase64()
```

## 配置

包含页面设置、元数据与元素默认值的完整配置：

```ts
const doc = createDocx({
  // Page setup
  page: {
    size: 'A4',
    orientation: 'portrait',
    margin: '20mm 25mm',
  },

  // Named styles
  styles: defineStyles({
    title: { fontSize: 24, fontWeight: 'bold' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),

  // Document metadata (appears in File → Info)
  metadata: {
    title: 'Annual Report 2026',
    creator: 'docx-kit',
    subject: 'Financial Report',
    description: 'Q1–Q4 financial performance summary',
    keywords: ['report', 'finance', '2026'],
  },

  // Element-level defaults
  defaults: {
    text: { fontFamily: 'Calibri', fontSize: 11 },
    paragraph: { textAlign: 'left' },
    table: { bordered: true },
    cell: { verticalAlign: 'middle' },
  },
})
```

## 后续学习

掌握以上内容后即可开始构建文档。按需选择下一步：

### 从现成样式开始

→ [样式预设](/zh-CN/guide/presets) — 开箱即用的 `classic`、`modern`、`academic` 配置

```ts
import { createDocx, modernPreset } from 'docx-kit'
const doc = createDocx(modernPreset.config)
```

### 统一多份文档的外观

→ [主题](/zh-CN/guide/themes) — `minimal`、`ocean`、`warm` 提供基于令牌的颜色与字体系统

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'
const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({ title: { color: '$colors.primary' } }),
})
```

### 让 AI/LLM 生成文档

→ [AI 模板](/zh-CN/guide/ai-templates) — 4 个内置模板与提示词构建器

```ts
import { reportTemplate } from 'docx-kit/ai'
const schema = reportTemplate.generate({ title: 'Q2 Report' /* ... */ })
```

### 通过 MCP 集成 Claude 或 ChatGPT

→ [MCP 服务端](/zh-CN/guide/mcp-server) — 面向 AI Agent 的 6 个工具与 1 个资源

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
// expose to MCP-compatible clients
```

### 学习完整构建器 API

→ [Builder API](/zh-CN/guide/builder-api) — 全部方法与选项

### 使用内置插件（二维码、图表、签名等）

→ [插件指南](/zh-CN/guide/plugins) — 22 个内置插件

## 详细阅读

- [Builder API](/zh-CN/guide/builder-api) — 学习全部链式 API 方法
- [CSS 风格样式](/zh-CN/guide/styling) — 掌握样式系统
- [表格](/zh-CN/guide/tables) — 创建数据表格
- [插件](/zh-CN/guide/plugins) — 通过自定义插件扩展
