# 快速開始

## 安裝

```bash
pnpm add docx-kit docx
```

請安裝 9.x 系列中不低於 9.15.0 的 `docx` 作為共用同儕相依套件。建構器、內建外掛與自訂外掛必須使用同一份 `docx` 實例。

### 可選同儕相依套件

部分內建外掛需要額外的套件：

```bash
# For QRCode plugin
pnpm add qrcode

# For ECharts plugin (browser only)
pnpm add echarts
```

## 快速開始

用 3 行程式碼建立第一個 .docx 文件：

```ts
import { createDocx } from 'docx-kit/node'

const doc = createDocx()

await doc
  .h1('Hello, World')
  .p('This is my first docx-kit document.')
  .save('hello.docx')
```

## 核心概念

docx-kit 圍繞三個核心概念建置：

| 概念 | 說明 |
| ----------- | --------------------------------------------------------------------------- |
| **樣式** | 將 CSS 風格規則（fontSize、color、margin 等）組織為樣式表 |
| **節點** | 內容建置區塊：標題、段落、表格、圖片和外掛 |
| **建構器** | 透過鏈式 API 組合節點並匯出為 .docx |

### 1. 定義樣式（CSS 風格）

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

### 2. 建置內容（鏈式 API）

```ts
const doc = createDocx({ styles })

await doc
  .h1('Quarterly Report', { className: 'title' })
  .p('Q1 2026 Financial Summary', { className: 'subtitle' })
  .p('Revenue grew 15% year-over-year...', { className: 'body' })
  .p('⚠️ Note: exchange rate fluctuations', { className: 'highlight' })
  .save('report.docx')
```

### 3. 匯出

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

## 設定

包含頁面設置、中繼資料與元素預設值的完整設定：

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

## 後續學習

掌握以上內容後即可開始建置文件。按需選擇下一步：

### 從現成樣式開始

→ [樣式預設](/zh-TW/guide/presets) — 開箱即用的 `classic`、`modern`、`academic` 設定

```ts
import { createDocx, modernPreset } from 'docx-kit'
const doc = createDocx(modernPreset.config)
```

### 統一多份文件的外觀

→ [主題](/zh-TW/guide/themes) — `minimal`、`ocean`、`warm` 提供基於令牌的顏色與字型系統

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'
const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({ title: { color: '$colors.primary' } }),
})
```

### 讓 AI/LLM 產生文件

→ [AI 範本](/zh-TW/guide/ai-templates) — 4 個內建範本與提示詞建構器

```ts
import { reportTemplate } from 'docx-kit/ai'
const schema = reportTemplate.generate({ title: 'Q2 Report' /* ... */ })
```

### 透過 MCP 整合 Claude 或 ChatGPT

→ [MCP 伺服器端](/zh-TW/guide/mcp-server) — 面向 AI Agent 的 6 個工具與 1 個資源

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
// expose to MCP-compatible clients
```

### 學習完整建構器 API

→ [Builder API](/zh-TW/guide/builder-api) — 全部方法與選項

### 使用內建外掛（二維碼、圖表、簽名等）

→ [外掛指南](/zh-TW/guide/plugins) — 22 個內建外掛

## 詳細閱讀

- [Builder API](/zh-TW/guide/builder-api) — 學習全部鏈式 API 方法
- [CSS 風格樣式](/zh-TW/guide/styling) — 掌握樣式系統
- [表格](/zh-TW/guide/tables) — 建立資料表格
- [外掛](/zh-TW/guide/plugins) — 透過自訂外掛擴充
