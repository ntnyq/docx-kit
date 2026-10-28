# 插件

docx-kit 当前提供 **22 个内置插件**，各自位于 `packages-plugins/` 下的独立工作区包中，通过 `.use(...)` 注册，再通过 `.plugin(name, options)` 调用。

## 插件目录

| 插件 | 包 | 节点名称 | 说明 |
| --- | --- | --- | --- |
| [徽章](./badge) | `@docxkit/plugin-badge` | `badge` | 状态标记与短标签 |
| [条形码](./barcode) | `@docxkit/plugin-barcode` | `barcode` | Code 128、EAN、UPC 等一维条形码 |
| [提示块](./callout) | `@docxkit/plugin-callout` | `callout` | 信息、警告、成功与危险提示框 |
| [更新日志](./changelog) | `@docxkit/plugin-changelog` | `changelog` | 带类型条目的发布说明表格 |
| [代码块](./code-block) | `@docxkit/plugin-code-block` | `codeBlock` | 等宽源码块，支持可选高亮 |
| [封面](./cover-page) | `@docxkit/plugin-cover-page` | `coverPage` | 专业标题页与报告封面 |
| [数据表格](./data-table) | `@docxkit/plugin-data-table` | `dataTable` | 带格式化辅助的对象数组表格 |
| [分隔线](./divider) | `@docxkit/plugin-divider` | `divider` | 横线与章节分隔 |
| [ECharts](./echarts) | `@docxkit/plugin-echarts` | `echarts` | 嵌入式 ECharts 图表 |
| [发票](./invoice) | `@docxkit/plugin-invoice` | `invoice` | 支持合计与税额计算的发票布局 |
| [信头](./letterhead) | `@docxkit/plugin-letterhead` | `letterhead` | 正式信函的公司品牌页头 |
| [会议纪要](./meeting-minutes) | `@docxkit/plugin-meeting-minutes` | `meetingMinutes` | 结构化会议摘要 |
| [页码](./page-number) | `@docxkit/plugin-page-number` | `pageNumber` | 当前页或第 X 页共 Y 页的域 |
| [属性表](./property-table) | `@docxkit/plugin-property-table` | `propertyTable` | 键值元数据表格 |
| [二维码](./qrcode) | `@docxkit/plugin-qrcode` | `qrcode` | 根据文字或 URL 生成二维码图片 |
| [任务清单](./checklist) | `@docxkit/plugin-checklist` | `checklist` | 分组 Word 原生复选框 |
| [插图](./figure) | `@docxkit/plugin-figure` | `figure` | 带题注与来源的图片 |
| [指标卡](./metric-card) | `@docxkit/plugin-metric-card` | `metricCard` | 多列指标摘要 |
| [签名区](./signature-block) | `@docxkit/plugin-signature-block` | `signatureBlock` | 审批与合同签名网格 |
| [时间线](./timeline) | `@docxkit/plugin-timeline` | `timeline` | 按时间顺序展示里程碑 |
| [目录](./toc) | `@docxkit/plugin-toc` | `toc` | 基于标题生成的 Word 目录域 |
| [水印](./watermark) | `@docxkit/plugin-watermark` | `watermark` | 简单文字水印 |

## 通用使用方式

```ts
import { createDocx, calloutPlugin } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'info',
    title: 'Heads up',
    content: 'This report is generated automatically.',
  })
```

## 兼容性说明

| 插件 | 额外依赖或运行时说明 |
| --- | --- |
| `barcode` | 需要 `bwip-js` 包 |
| `codeBlock` | 语法着色需要可选 `highlight.js` 对等依赖 |
| `echarts` | 需要 `echarts`，主要用于浏览器 DOM 环境 |
| `qrcode` | 需要 `qrcode` 包 |
| `toc` | 打开生成的 `.docx` 后，Word 通常会提示更新域 |
| `pageNumber` | 适合放在页眉或页脚 |

## 如何选择

- 需要常规文档流中的展示块时，使用 `callout`、`codeBlock`、`badge`、`divider` 等内容插件。
- 需要可复用报告结构时，使用 `coverPage`、`toc`、`pageNumber`、`letterhead` 等结构插件。
- 来源数据已有明确 schema 时，使用 `dataTable`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes` 等数据插件。
- 需要可扫描或可视化输出时，使用 `barcode`、`echarts`、`qrcode` 等媒体插件。

也支持通过 `definePlugin()` 创建自定义插件。详见[插件指南](/zh-CN/guide/plugins)与[包总览](/zh-CN/ecosystem/packages)。
