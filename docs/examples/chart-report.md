# Example: Chart Report

Generate an analytics report with embedded ECharts bar charts, pie charts, line charts, QR codes, and formatted tables.

## Full Code

```ts
import {
  createDocx,
  defineStyles,
  echartsPlugin,
  qrcodePlugin,
} from 'docx-kit'

// 1. Styles
const styles = defineStyles({
  title: {
    fontSize: 26,
    fontWeight: 'bold',
    color: '#1a1a2e',
    textAlign: 'center',
  },
  subtitle: {
    fontSize: 14,
    color: '#888',
    textAlign: 'center',
    marginBottom: 20,
  },
  h2: {
    fontSize: 18,
    fontWeight: 'bold',
    color: '#333',
    marginTop: 20,
    marginBottom: 10,
  },
  body: {
    fontSize: 12,
    lineHeight: 1.5,
  },
  captionStyle: {
    fontSize: 10,
    color: '#666',
    textAlign: 'center',
    fontStyle: 'italic',
  },
  tableHeader: {
    fontWeight: 'bold',
    backgroundColor: '#1e293b',
    color: '#ffffff',
    fontSize: 10,
  },
  tableCell: {
    fontSize: 10,
    verticalAlign: 'middle',
  },
})

// 2. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
  metadata: {
    title: 'Analytics Report — June 2026',
    creator: 'docx-kit',
    keywords: ['analytics', 'charts', 'report', '2026'],
  },
})

doc
  // Register plugins
  .use(echartsPlugin())
  .use(qrcodePlugin())

  // === Cover ===
  .h1('Analytics Dashboard', { className: 'title' })
  .p('June 2026', { className: 'subtitle' })

  .pageBreak()

  // === Section 1: Revenue Bar Chart ===
  .h2('1. Revenue by Product Line', { className: 'h2' })
  .p('The following chart shows revenue breakdown across our four main product lines for Q1 and Q2 2026.', { className: 'body' })

  .plugin('echarts', {
    option: {
      title: {
        text: 'Revenue by Product (Q1 vs Q2)',
        left: 'center',
        textStyle: { fontSize: 14 },
      },
      tooltip: { trigger: 'axis' },
      legend: {
        data: ['Q1 Revenue', 'Q2 Revenue'],
        bottom: 0,
      },
      xAxis: {
        type: 'category',
        data: ['Cloud SaaS', 'Enterprise', 'Mobile App', 'Consulting'],
      },
      yAxis: {
        type: 'value',
        axisLabel: { formatter: '${value}K' },
      },
      series: [
        {
          name: 'Q1 Revenue',
          type: 'bar',
          data: [420, 380, 250, 180],
          itemStyle: { color: '#93c5fd' },
        },
        {
          name: 'Q2 Revenue',
          type: 'bar',
          data: [480, 410, 320, 200],
          itemStyle: { color: '#2563eb' },
        },
      ],
    },
    width: 600,
    height: 360,
    caption: 'Figure 1: Revenue comparison by product line (in $K)',
  })

  .pageBreak()

  // === Section 2: Market Share Pie Chart ===
  .h2('2. Market Share Distribution', { className: 'h2' })
  .p('Market share analysis shows our Cloud SaaS product dominates at 42%.', { className: 'body' })

  .plugin('echarts', {
    option: {
      title: {
        text: 'Market Share by Product',
        left: 'center',
        textStyle: { fontSize: 14 },
      },
      tooltip: { formatter: '{b}: {c}%' },
      legend: {
        orient: 'vertical',
        right: 10,
        top: 'center',
      },
      series: [{
        type: 'pie',
        radius: ['40%', '70%'],
        center: ['40%', '50%'],
        data: [
          { name: 'Cloud SaaS', value: 42 },
          { name: 'Enterprise', value: 30 },
          { name: 'Mobile App', value: 18 },
          { name: 'Consulting', value: 10 },
        ],
        emphasis: {
          itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.3)' },
        },
        label: {
          formatter: '{b}: {d}%',
        },
      }],
    },
    width: 600,
    height: 360,
    caption: 'Figure 2: Current market share distribution',
  })

  .pageBreak()

  // === Section 3: User Growth Line Chart ===
  .h2('3. User Growth Trends', { className: 'h2' })
  .p('Active users have grown 152% over the past 12 months.', { className: 'body' })

  .plugin('echarts', {
    option: {
      title: {
        text: 'Monthly Active Users (12-Month Trend)',
        left: 'center',
        textStyle: { fontSize: 14 },
      },
      tooltip: { trigger: 'axis' },
      xAxis: {
        type: 'category',
        data: ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      },
      yAxis: {
        type: 'value',
        name: 'Users',
      },
      series: [
        {
          name: 'Active Users',
          type: 'line',
          data: [6200, 6800, 7500, 8100, 8800, 9100, 9800, 10500, 11500, 12800, 14200, 15600],
          smooth: true,
          lineStyle: { color: '#2563eb', width: 3 },
          itemStyle: { color: '#2563eb' },
          areaStyle: {
            color: {
              type: 'linear',
              x: 0, y: 0, x2: 0, y2: 1,
              colorStops: [
                { offset: 0, color: 'rgba(37,99,235,0.3)' },
                { offset: 1, color: 'rgba(37,99,235,0.02)' },
              ],
            },
          },
        },
        {
          name: 'New Signups',
          type: 'line',
          data: [250, 310, 420, 380, 450, 300, 520, 480, 620, 720, 780, 650],
          smooth: true,
          lineStyle: { color: '#22c55e', width: 2, type: 'dashed' },
          itemStyle: { color: '#22c55e' },
        },
      ],
    },
    width: 600,
    height: 380,
    caption: 'Figure 3: Monthly active users and new signups over 12 months',
  })

  .pageBreak()

  // === Section 4: Summary Table ===
  .h2('4. Key Performance Indicators', { className: 'h2' })

  interface KpiRow {
    metric: string
    current: string
    target: string
    status: string
  }

  doc.table<KpiRow>({
    columns: [
      { key: 'metric', title: 'KPI', width: '30%' },
      { key: 'current', title: 'Current', width: '25%', align: 'right' },
      { key: 'target', title: 'Target', width: '25%', align: 'right' },
      {
        key: 'status',
        title: 'Status',
        width: '20%',
        align: 'center',
        render: (val) => {
          const s = val as string
          const color = s === '✓ On Track' ? '#22c55e'
            : s === '⚠ At Risk' ? '#f59e0b'
            : '#ef4444'
          return [{ type: 'text', text: s, style: { color, fontWeight: 'bold' } }]
        },
      },
    ],
    data: [
      { metric: 'Monthly Revenue', current: '$480K', target: '$450K', status: '✓ On Track' },
      { metric: 'Active Users', current: '15,600', target: '14,000', status: '✓ On Track' },
      { metric: 'Customer Churn', current: '2.1%', target: '<2%', status: '⚠ At Risk' },
      { metric: 'NPS Score', current: '68', target: '75', status: '✗ Behind' },
      { metric: 'Response Time', current: '4.2h', target: '<3h', status: '✗ Behind' },
      { metric: 'Uptime', current: '99.97%', target: '99.9%', status: '✓ On Track' },
    ],
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

  // === Section 5: QR Code ===
  .h2('5. View Online Dashboard', { className: 'h2' })
  .p('Scan the QR code below to view the full interactive dashboard online:', {
    style: { textAlign: 'center', fontSize: 11 },
  })

  .plugin('qrcode', {
    text: 'https://analytics.example.com/dashboard/june-2026',
    size: 150,
    errorCorrectionLevel: 'H',
    caption: 'https://analytics.example.com',
  })

  // === Footer ===
  .pageBreak()
  .p('─'.repeat(80), { style: { fontSize: 6, color: '#ddd' } })
  .p('Generated by docx-kit on June 10, 2026', {
    style: { fontSize: 9, color: '#ccc', textAlign: 'center' },
  })
  .p('Data sourced from internal analytics. All figures are unaudited.', {
    style: { fontSize: 8, color: '#d4d4d4', textAlign: 'center' },
  })

// 3. Export in the browser
const blob = await doc.toBlob()
const url = URL.createObjectURL(blob)
const anchor = document.createElement('a')
anchor.href = url
anchor.download = 'analytics-report-june-2026.docx'
anchor.click()
URL.revokeObjectURL(url)
```

## What This Demonstrates

| Feature | Used In |
|---|---|
| `echartsPlugin` — bar chart | Revenue comparison (Q1 vs Q2) |
| `echartsPlugin` — donut/pie chart | Market share distribution |
| `echartsPlugin` — line chart with gradient area | User growth trend |
| `qrcodePlugin` | Link to online dashboard |
| Multiple `pageBreak()` | One chart per page |
| Typed table with color-coded `render()` | KPI status table |
| Rich ECharts options: tooltips, legends, gradients | All 3 charts |
| Chart captions with italic style | Below each chart |
| Document metadata | Title, keywords |
