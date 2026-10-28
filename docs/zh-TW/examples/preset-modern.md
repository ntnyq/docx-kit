# 範例：現代預設商務報告

使用 `modernPreset` 產生完整商務報告，採用 Calibri 字型與藍色強調標題，提供簡潔專業的樣式。

## 完整程式碼

```ts
import { createDocx, modernPreset, calloutPlugin, dataTablePlugin, pageNumberPlugin } from 'docx-kit/node'

const doc = createDocx({
  ...modernPreset.config,
  metadata: {
    title: 'Q1 2026 Business Report',
    creator: 'Finance Team',
    subject: 'Quarterly financial performance summary',
  },
})

doc
  // ── Title page ──
  .h1('Q1 2026 Business Report')
  .p('Quarterly financial performance summary', {
    style: { textAlign: 'center', color: '#64748b', fontSize: 12 },
  })
  .pageBreak()

  // ── Executive summary with callout ──
  .h2('Executive Summary')
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'success',
    title: 'On Track',
    content: 'Revenue exceeded projections by 12%. All key product lines showed strong growth.',
  })
  .p('This report summarizes Q1 2026 financial performance across all business segments. Key highlights include:')
  .bulletList([
    'Total revenue of $1.45M, up 12% YoY',
    'New customer acquisition up 23%',
    'Churn rate decreased by 0.5 percentage points',
    'NPS score improved to 72 (+5 points)',
  ])

  // ── Key metrics table ──
  .h2('Key Metrics')
  .use(dataTablePlugin())
  .plugin('dataTable', {
    data: [
      { metric: 'Total Revenue',     value: '$1.45M', change: '+12%' },
      { metric: 'New Customers',     value: '847',    change: '+23%' },
      { metric: 'Churn Rate',        value: '2.1%',   change: '-0.5%' },
      { metric: 'NPS Score',         value: '72',     change: '+5' },
    ],
    columns: [
      { key: 'metric', title: 'Metric',     align: 'left' },
      { key: 'value',  title: 'Value',      align: 'right' },
      { key: 'change', title: 'Change',     align: 'right' },
    ],
  })

  // ── Detailed financials ──
  .pageBreak()
  .h2('Financial Detail')

  .h3('Revenue by Segment')
  .table({
    columns: [
      { key: 'segment', title: 'Segment' },
      { key: 'q1_2026', title: 'Q1 2026', align: 'right' },
      { key: 'q1_2025', title: 'Q1 2025', align: 'right' },
      { key: 'growth',  title: 'YoY Growth', align: 'right' },
    ],
    data: [
      { segment: 'Enterprise', q1_2026: '$820K', q1_2025: '$680K', growth: '+21%' },
      { segment: 'SMB',        q1_2026: '$420K', q1_2025: '$390K', growth: '+8%'  },
      { segment: 'Consumer',   q1_2026: '$210K', q1_2025: '$220K', growth: '-5%'  },
    ],
    bordered: true,
    striped: true,
    headerCellStyle: { fontWeight: 'bold', backgroundColor: '#1B2A4A', color: '#ffffff' },
  })

  // ── Footer with page number ──
  .section({
    footer: {
      default: {
        children: [
          'Confidential — Q1 2026 Business Report',
          { type: 'plugin', name: 'pageNumber', options: { showTotal: true } },
        ],
      },
    },
  })
  .use(pageNumberPlugin())

await doc.save('q1-2026-report.docx')
```

## 預設提供的能力

`modernPreset` 自動提供：

- 全文使用 **Calibri** 字型
- **h1** 使用海軍藍 `#1B2A4A`、26pt、加粗，並帶 1.5pt 藍色下邊框
- **h2/h3** 使用藍色 `#2E75B6`
- **正文**使用 Calibri 11pt 與 1.5 倍行距
- **圖片居中**，垂直外邊距為 10pt

只有需要調整基本樣式時，才需要添加自訂樣式。

## 變體

### 自訂強調色

```ts
const doc = createDocx({
  ...modernPreset.config,
  styles: {
    ...modernPreset.config.styles,
    brandAccent: { color: '#8b5cf6', fontWeight: 'bold' },
  },
})

doc.p('Custom purple accent', { className: 'brandAccent' })
```

### 疊加主題

```ts
import { createDocx, modernPreset, useTheme, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  styles: {
    ...modernPreset.config.styles,
    highlight: {
      color: '$colors.accent',
      backgroundColor: '$colors.surface',
    },
  },
})
```

## 相關文件

- [樣式預設](/zh-TW/guide/presets) — 預設參考
- [主題](/zh-TW/guide/themes) — 主題系統
- [基本報告](/zh-TW/examples/basic-report) — 精簡建構器範例
- [圖表報告](/zh-TW/examples/chart-report) — 使用 ECharts 外掛
