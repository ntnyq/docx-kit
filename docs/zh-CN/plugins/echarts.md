# ECharts

把 ECharts 图表渲染成图片并嵌入文档。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `option` | `EChartsOption` | 图表配置，必填 |
| `width` | `number` | 图表宽度 |
| `height` | `number` | 图表高度 |
| `renderer` | `'canvas' \| 'svg'` | 渲染器 |
| `imageType` | `'png' \| 'svg'` | 输出图片格式 |
| `caption` | `string` | 图注 |

## 用法

```ts
createDocx().use(echartsPlugin()).plugin('echarts', {
  option: { xAxis: { data: ['Q1', 'Q2'] }, yAxis: {}, series: [{ type: 'bar', data: [12, 18] }] },
})
```

主要面向浏览器环境，使用前需要安装 `echarts`。
