# 外掛

外掛是 docx-kit 的主要擴充機制，無需修改核心編譯器即可添加新節點型別，同時保留統一的建置流程與型別安全的選項物件。

## 工作方式

每個內建或自訂外掛都遵循相同生命週期：

1. 匯入外掛工廠函式或外掛實例。
2. 使用 `.use(...)` 註冊。
3. 使用 `.plugin(name, options)` 呼叫。

```ts
import { calloutPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'success',
    title: 'Published',
    content: 'The release notes were exported successfully.',
  })
```

## 內建外掛範圍

目前工作區包含 **22 個內建外掛**：

- 內容區塊：`badge`、`callout`、`checklist`、`codeBlock`、`divider`、`watermark`
- 報告結構：`coverPage`、`letterhead`、`toc`、`pageNumber`
- 結構化業務資料：`dataTable`、`metricCard`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes`、`changelog`
- 媒體與嵌入：`barcode`、`echarts`、`figure`、`qrcode`
- 審批流程：`signatureBlock`

完整參考見[外掛章節](/zh-TW/plugins/)。

## 包結構

每個內建外掛都在 `packages-plugins/` 下作為獨立套件發佈：

```text
packages-plugins/
  badge/
  barcode/
  callout/
  changelog/
  checklist/
  code-block/
  cover-page/
  data-table/
  divider/
  echarts/
  figure/
  invoice/
  letterhead/
  meeting-minutes/
  metric-card/
  page-number/
  property-table/
  qrcode/
  signature-block/
  timeline/
  toc/
  watermark/
```

這樣可以獨立維護外掛的發佈、測試、同儕相依套件與文件。

## 同儕相依套件

部分外掛將體積較大的執行階段相依設為可選：

| 外掛 | 可選原因 |
| --- | --- |
| `@docxkit/plugin-barcode` | 僅產生條碼圖片時需要 `bwip-js` |
| `@docxkit/plugin-code-block` | 僅語法著色時需要 `highlight.js` |
| `@docxkit/plugin-echarts` | 僅使用圖表呈現時需要 `echarts` |
| `@docxkit/plugin-qrcode` | 僅產生二維碼圖片時需要 `qrcode` |

缺少可選同儕相依套件時，外掛會優雅降級，或拋出帶有解決建議的執行階段錯誤。

## 自訂外掛

需要業務或文件流程專屬節點型別時，使用 `definePlugin()`。

```ts
import { definePlugin } from 'docx-kit'
import { Paragraph } from 'docx'

interface DisclaimerOptions {
  text: string
}

const disclaimerPlugin = definePlugin<'disclaimer', DisclaimerOptions>({
  name: 'disclaimer',
  render(options) {
    return new Paragraph(options.text)
  },
})
```

好的自訂外掛應當：

- 接收小而明確的選項物件
- 回傳常規 `docx` 節點或節點陣列
- 在呈現函式附近完成驗證
- 避免隱藏無關的文件修改

## 外掛配套工具

- `@docxkit/create-plugin`：建立外掛套件鷹架
- `@docxkit/pdk`：外掛測試輔助工具
- `@docxkit/loader`：從 inline、npm、URL 或本地來源載入外掛
- `@docxkit/registry`：發現帶 docx-kit 關鍵字的外掛
- `@docxkit/mcp`：向 AI Agent 提供支援外掛的文件產生工具

各部分所在位置見 [Monorepo 結構](/zh-TW/ecosystem/monorepo)與[套件總覽](/zh-TW/ecosystem/packages)。
