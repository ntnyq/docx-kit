# 外掛

docx-kit 目前提供 **22 個內建外掛**，各自位於 `packages-plugins/` 下的獨立工作區包中，透過 `.use(...)` 註冊，再透過 `.plugin(name, options)` 呼叫。

## 外掛目錄

| 外掛 | 套件 | 節點名稱 | 說明 |
| --- | --- | --- | --- |
| [徽章](./badge) | `@docxkit/plugin-badge` | `badge` | 狀態標記與短標籤 |
| [條碼](./barcode) | `@docxkit/plugin-barcode` | `barcode` | Code 128、EAN、UPC 等一維條碼 |
| [提示區塊](./callout) | `@docxkit/plugin-callout` | `callout` | 資訊、警告、成功與危險提示框 |
| [更新記錄](./changelog) | `@docxkit/plugin-changelog` | `changelog` | 帶型別條目的發佈說明表格 |
| [程式碼區塊](./code-block) | `@docxkit/plugin-code-block` | `codeBlock` | 等寬原始碼區塊，支援可選高亮 |
| [封面](./cover-page) | `@docxkit/plugin-cover-page` | `coverPage` | 專業標題頁與報告封面 |
| [資料表格](./data-table) | `@docxkit/plugin-data-table` | `dataTable` | 帶格式化輔助的物件陣列表格 |
| [分隔線](./divider) | `@docxkit/plugin-divider` | `divider` | 橫線與章節分隔 |
| [ECharts](./echarts) | `@docxkit/plugin-echarts` | `echarts` | 嵌入式 ECharts 圖表 |
| [發票](./invoice) | `@docxkit/plugin-invoice` | `invoice` | 支援合計與稅額計算的發票佈局 |
| [信頭](./letterhead) | `@docxkit/plugin-letterhead` | `letterhead` | 正式信函的公司品牌頁頭 |
| [會議記錄](./meeting-minutes) | `@docxkit/plugin-meeting-minutes` | `meetingMinutes` | 結構化會議摘要 |
| [頁碼](./page-number) | `@docxkit/plugin-page-number` | `pageNumber` | 目前頁或第 X 頁共 Y 頁的域 |
| [屬性表](./property-table) | `@docxkit/plugin-property-table` | `propertyTable` | 鍵值中繼資料表格 |
| [二維碼](./qrcode) | `@docxkit/plugin-qrcode` | `qrcode` | 根據文字或 URL 產生二維碼圖片 |
| [工作清單](./checklist) | `@docxkit/plugin-checklist` | `checklist` | 分組工作，支援圖形與 Word 原生核取方塊 |
| [插圖](./figure) | `@docxkit/plugin-figure` | `figure` | 帶題注與來源的圖片 |
| [指標卡](./metric-card) | `@docxkit/plugin-metric-card` | `metricCard` | 多欄指標摘要 |
| [簽名區](./signature-block) | `@docxkit/plugin-signature-block` | `signatureBlock` | 審批與合同簽名網格 |
| [時間軸](./timeline) | `@docxkit/plugin-timeline` | `timeline` | 按時間順序展示里程碑 |
| [目錄](./toc) | `@docxkit/plugin-toc` | `toc` | 基於標題產生的 Word 目錄域 |
| [浮水印](./watermark) | `@docxkit/plugin-watermark` | `watermark` | 簡單文字浮水印 |

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

## 兼容性說明

| 外掛 | 額外相依或執行階段說明 |
| --- | --- |
| `barcode` | 需要 `bwip-js` 包 |
| `codeBlock` | 語法著色需要可選 `highlight.js` 同儕相依套件 |
| `echarts` | 需要 `echarts`，主要用於瀏覽器 DOM 環境 |
| `qrcode` | 需要 `qrcode` 包 |
| `toc` | 打開產生的 `.docx` 後，Word 通常會提示更新域 |
| `pageNumber` | 適合放在頁眉或頁腳 |

## 如何選擇

- 需要常規文件流中的展示區塊時，使用 `callout`、`codeBlock`、`badge`、`divider` 等內容外掛。
- 需要可重用報告結構時，使用 `coverPage`、`toc`、`pageNumber`、`letterhead` 等結構外掛。
- 來源資料已有明確 schema 時，使用 `dataTable`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes` 等資料外掛。
- 需要可掃描或可視化輸出時，使用 `barcode`、`echarts`、`qrcode` 等媒體外掛。

也支援透過 `definePlugin()` 建立自訂外掛。詳見[外掛指南](/zh-TW/guide/plugins)與[套件總覽](/zh-TW/ecosystem/packages)。
