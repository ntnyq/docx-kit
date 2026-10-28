# 套件總覽

本頁按套件展示目前 docx-kit monorepo 的執行階段、工具、預設和主題，與外掛參考文件互補。

## 核心執行階段套件

| 套件 | 用途 | 使用場景 |
| --- | --- | --- |
| `docx-kit` | 整合套件 | 希望統一匯入 API 的應用開發者的預設選擇 |
| `@docxkit/core` | 核心建構器、編譯器、DSL、樣式與工具 | 需要精簡執行階段或建置高級整合 |
| `@docxkit/types` | 共享型別定義 | 僅需要型別的套件或應用 |
| `@docxkit/markdown` | 將 Markdown 轉換為文件 DSL | 使用 Markdown 輸入重用文件樣式、主題與外掛 |
| `@docxkit/renderer` | 瀏覽器預覽呈現 | 需要瀏覽器內 `.docx` 預覽或嵌入 Office Online |

## 擴充與工具

| 套件 | 用途 | 核心能力 |
| --- | --- | --- |
| `@docxkit/loader` | 動態載入外掛 | 支援 inline、npm、URL 與本地外掛來源 |
| `@docxkit/pdk` | 外掛開發工具套件 | 外掛呈現測試與可重用測試上下文 |
| `@docxkit/registry` | 外掛登錄服務搜尋 | 發現為 docx-kit 發佈的外掛 |
| `@docxkit/create-plugin` | 外掛鷹架 | 按約定目錄結構建立外掛套件 |

## AI 與 Agent 套件

| 套件 | 用途 | 核心能力 |
| --- | --- | --- |
| `@docxkit/ai` | 適合 AI 的範本系統 | 報告、發票、履歷、信函範本與提示詞建置 |
| `@docxkit/mcp` | MCP 伺服器端 | 向 AI Agent 和助手提供文件產生工具 |

## 內建外掛套件

| 套件 | 功能 |
| --- | --- |
| `@docxkit/plugin-badge` | 狀態標籤 |
| `@docxkit/plugin-barcode` | 一維條碼 |
| `@docxkit/plugin-callout` | 提示區塊 |
| `@docxkit/plugin-changelog` | 更新記錄表格 |
| `@docxkit/plugin-code-block` | 原始碼排版 |
| `@docxkit/plugin-cover-page` | 報告與文件封面 |
| `@docxkit/plugin-data-table` | 結構化物件陣列表格 |
| `@docxkit/plugin-divider` | 章節分隔線 |
| `@docxkit/plugin-echarts` | 圖表呈現 |
| `@docxkit/plugin-invoice` | 賬單文件 |
| `@docxkit/plugin-letterhead` | 品牌信頭 |
| `@docxkit/plugin-meeting-minutes` | 會議記錄 |
| `@docxkit/plugin-page-number` | 頁碼域 |
| `@docxkit/plugin-property-table` | 鍵值表格 |
| `@docxkit/plugin-qrcode` | 二維碼 |
| `@docxkit/plugin-checklist` | 分組 Word 原生核取方塊 |
| `@docxkit/plugin-figure` | 帶題注與來源的圖片 |
| `@docxkit/plugin-metric-card` | 多欄指標摘要 |
| `@docxkit/plugin-signature-block` | 審批與簽名佈局 |
| `@docxkit/plugin-timeline` | 里程碑時間軸 |
| `@docxkit/plugin-toc` | 目錄域 |
| `@docxkit/plugin-watermark` | 浮水印文字 |

每個外掛的完整選項參考見[外掛章節](/zh-TW/plugins/)。

## 預設套件

| 套件 | 樣式方向 |
| --- | --- |
| `@docxkit/preset-classic` | 正式公文樣式 |
| `@docxkit/preset-modern` | 適合商務文件的預設樣式 |
| `@docxkit/preset-academic` | 論文與學術寫作排版 |

## 主題套件

| 套件 | 主題方向 |
| --- | --- |
| `@docxkit/theme-minimal` | 中性簡約配色 |
| `@docxkit/theme-ocean` | 藍色主導的配色系統 |
| `@docxkit/theme-warm` | 琥珀色與暖色配色系統 |

## 推薦入口

- 應用開發選擇 `docx-kit`。
- 適配器、外掛或受限執行階段套件選擇 `@docxkit/core`。
- 瀏覽器預覽選擇 `@docxkit/renderer`。
- 擴充生態選擇 `@docxkit/pdk` 與 `@docxkit/create-plugin`。
- AI 驅動的文件流程選擇 `@docxkit/ai` 或 `@docxkit/mcp`。
