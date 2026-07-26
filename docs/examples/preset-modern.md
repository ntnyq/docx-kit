# Example: Modern Preset Business Report

A complete business report using the `modernPreset` for clean, professional styling with Calibri throughout and blue accent headings.

## Full Code

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

## What the Preset Gives You

The `modernPreset` automatically provides:

- **Calibri** font throughout the document
- **h1** in navy `#1B2A4A`, 26pt, bold, with a 1.5pt blue underline
- **h2/h3** in blue `#2E75B6`
- **Body text** in Calibri 11pt with 1.5× line height
- **Centered images** with 10pt vertical margin

You only need to add custom styles when you want to deviate from this base.

## Variations

### Add a Custom Accent Color

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

### Add a Theme on Top

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

## See Also

- [Style Presets](/guide/presets) — Preset reference
- [Themes](/guide/themes) — Theme system
- [Basic Report](/examples/basic-report) — Minimal builder example
- [Chart Report](/examples/chart-report) — With ECharts plugin
