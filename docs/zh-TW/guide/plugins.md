# 外掛系統

docx-kit 透過 `.use(...)` 註冊外掛，透過 `.plugin(name, options)` 呼叫外掛。

## 基本模式

```ts
import { calloutPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'info',
    title: '提示',
    content: '這是一段由外掛生成的內容。',
  })
```

## 目前內建外掛

目前 workspace 共有 **18 個內建外掛**，而且每個外掛都有獨立 package 與獨立文件頁。

- 內容區塊：`badge`、`callout`、`codeBlock`、`divider`、`watermark`
- 文件結構：`coverPage`、`letterhead`、`toc`、`pageNumber`
- 結構化資料：`dataTable`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes`、`changelog`
- 媒體與嵌入：`echarts`、`qrcode`
- 簽署流程：`signatureBlock`

完整列表見 [外掛總覽](/zh-TW/plugins/)。
