# ECharts

Renders ECharts charts as embedded images in your Word documents. Supports all ECharts chart types — bar, line, pie, scatter, radar, heatmap, and more.

## Import

```ts
import { echartsPlugin, type EChartsPluginOptions } from 'docx-kit'
```

## Dependencies

Requires `echarts` as a peer dependency:

```bash
pnpm add echarts
```

> **Platform Note:** ECharts rendering requires a browser DOM environment. In Node.js, it throws an error prompting you to provide a server-side canvas solution.

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `option` | `EChartsOption` | _(required)_ | Full ECharts option object (series, axes, title, etc.) |
| `width` | `number` | `640` | Chart width in pixels |
| `height` | `number` | `360` | Chart height in pixels |
| `renderer` | `'canvas' \| 'svg'` | `'canvas'` | ECharts rendering engine |
| `imageType` | `'png' \| 'svg'` | `'png'` | Output image format |
| `caption` | `string` | — | Optional caption below the chart |

## Examples

### Bar Chart

```ts
import { createDocx, echartsPlugin } from 'docx-kit'

const doc = createDocx()
  .use(echartsPlugin())
  .h1('Monthly Revenue')
  .plugin('echarts', {
    option: {
      title: { text: 'Monthly Revenue 2026' },
      tooltip: {},
      xAxis: {
        type: 'category',
        data: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      },
      yAxis: { type: 'value' },
      series: [{
        name: 'Revenue',
        type: 'bar',
        data: [120, 200, 150, 80, 70, 110],
        itemStyle: { color: '#2563eb' },
      }],
    },
    width: 640,
    height: 360,
    caption: 'Figure 1: Monthly revenue by product line',
  })
  .save('bar-chart.docx')
```

### Line Chart

```ts
const doc = createDocx()
  .use(echartsPlugin())
  .h1('User Growth')
  .plugin('echarts', {
    option: {
      title: { text: 'User Growth Trend' },
      tooltip: { trigger: 'axis' },
      legend: { data: ['Users', 'Active'] },
      xAxis: { type: 'category', data: ['Jan', 'Feb', 'Mar', 'Apr', 'May'] },
      yAxis: { type: 'value' },
      series: [
        {
          name: 'Users',
          type: 'line',
          data: [150, 230, 224, 218, 335],
          smooth: true,
          itemStyle: { color: '#2563eb' },
        },
        {
          name: 'Active',
          type: 'line',
          data: [120, 182, 191, 234, 290],
          smooth: true,
          itemStyle: { color: '#16a34a' },
        },
      ],
    },
    caption: 'Figure 1: Monthly user growth and active users',
  })
  .save('line-chart.docx')
```

### Pie Chart

```ts
const doc = createDocx()
  .use(echartsPlugin())
  .h1('Market Share Analysis')
  .plugin('echarts', {
    option: {
      title: { text: 'Revenue by Product' },
      tooltip: { trigger: 'item' },
      legend: { orient: 'vertical', left: 'left' },
      series: [{
        name: 'Revenue Share',
        type: 'pie',
        radius: '60%',
        data: [
          { name: 'Product A', value: 1048 },
          { name: 'Product B', value: 735 },
          { name: 'Product C', value: 580 },
          { name: 'Product D', value: 484 },
          { name: 'Product E', value: 300 },
        ],
        emphasis: {
          itemStyle: {
            shadowBlur: 10,
            shadowOffsetX: 0,
            shadowColor: 'rgba(0, 0, 0, 0.5)',
          },
        },
      }],
    },
    caption: 'Figure 2: Revenue distribution across product lines',
  })
  .save('pie-chart.docx')
```

### Scatter Chart

```ts
const doc = createDocx()
  .use(echartsPlugin())
  .h1('Correlation Analysis')
  .plugin('echarts', {
    option: {
      title: { text: 'Height vs Weight Correlation' },
      xAxis: { name: 'Height (cm)', type: 'value' },
      yAxis: { name: 'Weight (kg)', type: 'value' },
      series: [{
        type: 'scatter',
        data: [
          [160, 55], [165, 60], [170, 65], [175, 70],
          [180, 75], [185, 80], [190, 85], [155, 50],
          [162, 58], [168, 63], [172, 68], [178, 72],
        ],
        itemStyle: { color: '#2563eb' },
      }],
    },
    caption: 'Figure 3: Height vs weight scatter plot (n = 12)',
  })
  .save('scatter-chart.docx')
```

### Multi-Chart Report

```ts
const doc = createDocx()
  .use(echartsPlugin())
  .h1('Q2 Analytics Report')

  // Bar chart — revenue
  .h2('Revenue by Quarter')
  .plugin('echarts', {
    option: {
      title: { text: 'Quarterly Revenue' },
      xAxis: { type: 'category', data: ['Q1', 'Q2', 'Q3', 'Q4'] },
      yAxis: { type: 'value' },
      series: [{
        type: 'bar',
        data: [820, 932, 901, 1347],
        itemStyle: { color: '#2563eb' },
      }],
    },
    caption: 'Figure 1: Quarterly revenue (in thousands)',
  })

  .pageBreak()

  // Line chart — growth
  .h2('Monthly Growth Rate')
  .plugin('echarts', {
    option: {
      title: { text: 'Monthly Growth Rate' },
      xAxis: { type: 'category', data: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'] },
      yAxis: { type: 'value', axisLabel: { formatter: '{value}%' } },
      series: [{
        type: 'line',
        data: [5.2, 3.8, 7.1, 4.9, 6.3, 8.5],
        smooth: true,
        areaStyle: { color: 'rgba(37, 99, 235, 0.1)' },
        itemStyle: { color: '#2563eb' },
      }],
    },
    caption: 'Figure 2: Month-over-month growth rate',
  })

  .pageBreak()

  // Pie chart — distribution
  .h2('Revenue by Region')
  .plugin('echarts', {
    option: {
      title: { text: 'Revenue by Region' },
      series: [{
        type: 'pie',
        radius: ['40%', '70%'],
        data: [
          { name: 'North America', value: 45 },
          { name: 'Europe', value: 25 },
          { name: 'Asia Pacific', value: 20 },
          { name: 'Others', value: 10 },
        ],
      }],
    },
    caption: 'Figure 3: Regional revenue distribution',
  })

  .save('analytics-report.docx')
```

### Custom Dimensions

```ts
const doc = createDocx()
  .use(echartsPlugin())
  .h1('Wide Chart')
  .plugin('echarts', {
    option: {
      title: { text: '12-Month Trend' },
      xAxis: {
        type: 'category',
        data: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      },
      yAxis: { type: 'value' },
      series: [{
        type: 'line',
        data: [100, 120, 115, 140, 160, 180, 200, 190, 210, 230, 220, 250],
      }],
    },
    width: 800,
    height: 300,
  })
  .save('wide-chart.docx')
```
