# 提示區塊（Callout）

帶圖標的彩色資訊、警告、成功和危險提示框。

## 匯入

```ts
import { calloutPlugin, type CalloutOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `content` | `string` | _必填_ | 提示區塊正文 |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | _必填_ | 提示區塊樣式，控制圖標、顏色與色調 |
| `title` | `string` | — | 正文前的可選加粗標題 |
| `styles` | `CalloutStyles` | — | 本次呼叫的部位樣式 |

### 型別預設

| 型別 | 圖標 | 背景色 | 邊框顏色 |
|------|------|------------|--------------|
| `info` | ℹ️ | `#D6E4F0` | `#4472C4` |
| `warning` | ⚠️ | `#FFF2CC` | `#FFC000` |
| `success` | ✅ | `#E2F0D9` | `#70AD47` |
| `danger` | ⚠️ | `#FCE4D6` | `#FF0000` |

## 範例

### 資訊提示

```ts
import { createDocx, calloutPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(calloutPlugin())
  .h1('System Maintenance')
  .plugin('callout', {
    type: 'info',
    content: 'The system will be upgraded tonight at 22:00. No downtime is expected.',
  })
  .save('callout.docx')
```

### 帶標題的警告

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Important Notice')
  .plugin('callout', {
    type: 'warning',
    title: '注意',
    content: '此操作不可撤销，请确认后再提交。',
  })
  .save('warning.docx')
```

### 成功確認

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Deployment Status')
  .plugin('callout', {
    type: 'success',
    title: 'Deployment Successful',
    content: 'Version 2.4.1 has been deployed to production. All health checks passed.',
  })
  .save('success.docx')
```

### 危險提醒

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Security Alert')
  .plugin('callout', {
    type: 'danger',
    title: 'Critical Vulnerability',
    content: 'A critical security vulnerability (CVE-2026-XXXX) has been detected. Immediate action required.',
  })
  .save('danger.docx')
```

### 多個提示區塊

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Weekly Summary')

  .plugin('callout', {
    type: 'success',
    title: 'Completed',
    content: 'Sprint goals met on time. 12 features shipped.',
  })

  .plugin('callout', {
    type: 'warning',
    title: 'At Risk',
    content: 'API rate limiting needs attention before next sprint.',
  })

  .plugin('callout', {
    type: 'info',
    title: 'Upcoming',
    content: 'Team retro scheduled for Friday at 3 PM.',
  })

  .save('multiple-callouts.docx')
```

## 樣式自訂

傳入 `{ styles }` 給外掛工廠，可設定目前外掛實例的預設樣式；透過 `options.styles` 覆寫單次呼叫。外掛套件和 `docx-kit` 的 Node/browser 入口均匯出 `CalloutConfig` / `CalloutStyles` 與 `DataTableConfig` / `DataTableStyles`。

優先順序由低至高：**內建預設值（含佈景主題值）→ 初始化 `styles` → 呼叫通用 `style`（`.plugin()` 第三個參數）→ 呼叫部位 `styles`**。省略或 `undefined` 欄位沿用預設值；巢狀普通物件逐欄位合併，陣列整體取代，`0` 和 `false` 均生效。每次呼叫獨立解析，不修改設定物件。單位和 `$colors.primary`、`$fontSize.lg`、`$spacing.md` 等主題 token 沿用現有樣式編譯器。

| 部位 | 樣式型別 | 作用範圍 |
| --- | --- | --- |
| `cell` | `Omit<CellStyleRule, 'height'>` | 容器背景、框線、寬度與內距 |
| `paragraph` | `ParagraphStyleRule` | 共用段落的對齊、間距與分頁控制 |
| `title` | `TextStyleRule` | 標題文字 |
| `content` | `TextStyleRule` | 內文文字 |

```ts
import { calloutPlugin, createDocx } from 'docx-kit/node'

const doc = createDocx({
  theme: { colors: { primary: '#2563EB' }, spacing: { md: 8 } },
})
  .use(calloutPlugin({
    styles: {
      cell: { margin: '8pt 10pt', borderLeft: { width: 3 } },
      title: { fontSize: 16, color: '$colors.primary' },
      content: { fontSize: 11 },
    },
  }))
  .plugin('callout', {
    type: 'info',
    title: 'Note',
    content: 'Custom body text',
    styles: {
      cell: { marginBottom: 0 },
      title: { bold: false, color: '#DC2626' },
      paragraph: { lineHeight: 1.5 },
    },
  }, { fontFamily: 'Arial' })
```

### JSON DSL、AI 與 MCP

JSON DSL 在呼叫的 `options.styles` 內填寫與 Builder 相同的可序列化部位樣式。AI 工具定義和 schema 指南已包含這些選項；MCP 透過 `docx-kit://schema` 和 `get_plugin_help().optionsSchema` 提供對應結構。`validate_schema` 會拒絕未知部位、放錯部位的屬性及錯誤的樣式值。其他內建外掛也提供對應的部位樣式契約。

```json
{
  "type": "plugin",
  "name": "callout",
  "options": {
    "type": "info",
    "title": "Note",
    "content": "Body",
    "styles": { "title": { "color": "#2563EB" }, "cell": { "margin": 0 } }
  }
}
```
