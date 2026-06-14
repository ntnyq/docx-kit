# ECharts

把 ECharts 圖表渲染成圖片並嵌入文件。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `option` | `EChartsOption` | 圖表設定，必填 |
| `width` | `number` | 圖表寬度 |
| `height` | `number` | 圖表高度 |
| `renderer` | `'canvas' \| 'svg'` | 渲染器 |
| `imageType` | `'png' \| 'svg'` | 輸出圖片格式 |
| `caption` | `string` | 圖註 |

## 用法

```ts
createDocx().use(echartsPlugin()).plugin('echarts', {
  option: { xAxis: { data: ['Q1', 'Q2'] }, yAxis: {}, series: [{ type: 'bar', data: [12, 18] }] },
})
```

主要面向瀏覽器環境，使用前需要安裝 `echarts`。
