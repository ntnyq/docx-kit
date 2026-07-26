# 插件系统

docx-kit 通过 `.use(...)` 注册插件，通过 `.plugin(name, options)` 调用插件。

## 基本模式

```ts
import { calloutPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'info',
    title: '提示',
    content: '这是一段由插件生成的内容。',
  })
```

## 当前内置插件

当前 workspace 内共有 **19 个内置插件**，并且每个插件都拥有独立包与独立文档页。

- 内容块：`badge`、`callout`、`codeBlock`、`divider`、`watermark`
- 文档结构：`coverPage`、`letterhead`、`toc`、`pageNumber`
- 结构化数据：`dataTable`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes`、`changelog`
- 媒体与嵌入：`barcode`、`echarts`、`qrcode`
- 审批签署：`signatureBlock`

完整列表见 [插件总览](/zh-CN/plugins/)。
