# Example: Basic Report

A complete example generating a multi-page business report with headings, formatted text, tables, page breaks, and a QR code.

## Full Code

```ts
import { createDocx, defineStyles, qrcodePlugin } from 'docx-kit'

// 1. Define styles
const styles = defineStyles({
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#1a1a2e',
    textAlign: 'center',
    marginBottom: 5,
  },
  subtitle: {
    fontSize: 16,
    color: '#666',
    textAlign: 'center',
    marginBottom: 20,
  },
  h2: {
    fontSize: 18,
    fontWeight: 'bold',
    color: '#333',
    marginTop: 15,
    marginBottom: 8,
  },
  h3: {
    fontSize: 14,
    fontWeight: 'bold',
    color: '#555',
    marginTop: 10,
    marginBottom: 5,
  },
  body: {
    fontSize: 12,
    lineHeight: 1.5,
    fontFamily: 'Calibri',
  },
  highlight: {
    backgroundColor: '#fff3cd',
    border: {
      style: 'single',
      width: 1,
      color: '#ffc107',
    },
    marginLeft: 20,
    marginRight: 20,
  },
  tableHeader: {
    fontWeight: 'bold',
    backgroundColor: '#1e293b',
    color: '#ffffff',
    fontSize: 10,
  },
  tableCell: {
    fontSize: 10,
  },
  small: {
    fontSize: 9,
    color: '#999',
  },
  footer: {
    fontSize: 9,
    color: '#999',
    textAlign: 'center',
    marginTop: 30,
  },
})

// 2. Data
interface MetricRow {
  metric: string
  q1: string
  q2: string
  change: string
}

const metrics: MetricRow[] = [
  { metric: 'Revenue', q1: '$1.20M', q2: '$1.45M', change: '+20.8%' },
  { metric: 'Costs', q1: '$0.85M', q2: '$0.92M', change: '+8.2%' },
  { metric: 'Profit', q1: '$0.35M', q2: '$0.53M', change: '+51.4%' },
  { metric: 'Active Users', q1: '12,450', q2: '15,320', change: '+23.0%' },
]

// 3. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
  metadata: {
    title: 'Q2 2026 Business Report',
    creator: 'docx-kit',
    subject: 'Quarterly Report',
    keywords: ['report', 'q2', '2026'],
  },
})

doc
  // === Cover ===
  .h1('Q2 2026 Business Report', { className: 'title' })
  .p('June 30, 2026', { className: 'subtitle' })
  .p('Prepared by: Finance Department', {
    style: { textAlign: 'center', fontSize: 11, color: '#888' },
  })

  .pageBreak()

  // === Executive Summary ===
  .h2('1. Executive Summary', { className: 'h2' })
  .p(
    'The second quarter of 2026 showed strong performance across all key metrics. '
    + 'Revenue exceeded projections by 12%, driven by new product launches and '
    + 'expansion into international markets.',
    { className: 'body' },
  )
  .p(
    '⚠️ Key Highlight: Revenue growth accelerated to 20.8% YoY, the highest '
    + 'quarterly growth rate in company history.',
    { className: 'highlight' },
  )

  // === Financial Metrics ===
  .h2('2. Financial Metrics', { className: 'h2' })
  .p('The following table summarizes Q1 vs Q2 performance:', { className: 'body' })

  .table<MetricRow>({
    columns: [
      { key: 'metric', title: 'Metric', width: '25%' },
      { key: 'q1', title: 'Q1 2026', width: '25%', align: 'right' },
      { key: 'q2', title: 'Q2 2026', width: '25%', align: 'right' },
      {
        key: 'change',
        title: 'Change',
        width: '25%',
        align: 'center',
        render: (val) => [
          {
            type: 'text',
            text: val as string,
            style: { color: '#22c55e', fontWeight: 'bold' },
          },
        ],
      },
    ],
    data: metrics,
    bordered: true,
    striped: true,
    headerCellStyle: {
      fontWeight: 'bold',
      backgroundColor: '#1e293b',
      color: '#ffffff',
      fontSize: 10,
    },
    cellStyle: {
      fontSize: 10,
      verticalAlign: 'middle',
    },
  })

  // === Operational Metrics ===
  .h2('3. Operational Highlights', { className: 'h2' })

  .h3('3.1 Product Development', { className: 'h3' })
  .p(
    'The engineering team shipped 3 major features in Q2, including the new '
    + 'analytics dashboard and real-time collaboration tools.',
    { className: 'body' },
  )

  .h3('3.2 Customer Success', { className: 'h3' })
  .p(
    'Customer satisfaction score (CSAT) improved from 87 to 91. Net Promoter Score '
    + '(NPS) reached 72, placing us in the top quartile of SaaS companies.',
    { className: 'body' },
  )

  .h3('3.3 Team Growth', { className: 'h3' })
  .p(
    'Headcount grew from 85 to 102 employees. Engineering team expanded by 8, '
    + 'Sales by 5, and Marketing by 4.',
    { className: 'body' },
  )

  .pageBreak()

  // === Outlook ===
  .h2('4. Q3 Outlook', { className: 'h2' })
  .p(
    'We project Q3 revenue of $1.65M–$1.75M, driven by the upcoming mobile app '
    + 'launch and enterprise partnership pipeline. Key risks include:',
    { className: 'body' },
  )
  .p('• Exchange rate fluctuations in European markets', { className: 'body' })
  .p('• Supply chain delays for hardware components', { className: 'body' })
  .p('• Competitive pressure in the APAC region', { className: 'body' })

  .h2('5. Contact Information', { className: 'h2' })
  .use(qrcodePlugin)
  .plugin('qrcode', {
    text: 'https://example.com/investor-relations',
    size: 150,
    caption: 'Scan to access our investor portal',
  })

  .pageBreak()

  // === Footer ===
  .p('Confidential — For Internal Use Only', { className: 'footer' })
  .p('© 2026 Example Corp. All rights reserved.', { className: 'small' })

  // 4. Export
  .save('q2-2026-business-report.docx')
```

## What This Demonstrates

| Feature | Used In |
|---|---|
| `defineStyles()` with class names | Headings, body text, highlights, footer |
| `className` on nodes | All headings & paragraphs |
| Inline `style:` overrides | Cover date text |
| Headings (h1, h2, h3) | Document structure |
| Typed tables + custom renderers | Financial metrics table |
| `striped: true`, `bordered: true` | Table styling |
| `pageBreak()` between sections | Cover → content → outlook → footer |
| Built-in `qrcodePlugin()` | Contact section |
| Page config + metadata | Document setup |
