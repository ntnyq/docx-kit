# 套件總覽

這個頁面按照 monorepo 最新狀態整理所有公開套件。

## 基礎執行時

| 套件 | 說明 |
| --- | --- |
| `docx-kit` | 預設總入口 |
| `@docxkit/core` | 核心文件建構與編譯能力 |
| `@docxkit/types` | 共用型別 |
| `@docxkit/renderer` | 瀏覽器內 `.docx` 預覽 |

## 擴充與工具

| 套件 | 說明 |
| --- | --- |
| `@docxkit/loader` | 從 inline、npm、URL、本地載入外掛 |
| `@docxkit/pdk` | 自訂外掛測試輔助 |
| `@docxkit/registry` | 基於 npm 關鍵字的外掛探索 |
| `@docxkit/create-plugin` | 新外掛腳手架 |

## AI 與 Agent

| 套件 | 說明 |
| --- | --- |
| `@docxkit/ai` | report、invoice、resume、letter 模板與 prompt 工具 |
| `@docxkit/mcp` | 面向 AI Agent 的 MCP 介面 |

## 預設與主題

- 預設：`@docxkit/preset-classic`、`@docxkit/preset-modern`、`@docxkit/preset-academic`
- 主題：`@docxkit/theme-minimal`、`@docxkit/theme-ocean`、`@docxkit/theme-warm`

## 外掛套件

所有內建外掛套件都在 [外掛總覽](/zh-TW/plugins/) 中逐一列出，並附上獨立文件頁。
