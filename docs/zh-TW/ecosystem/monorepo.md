# Monorepo 結構

倉庫使用 pnpm workspace monorepo，目前文件已依照最新結構覆蓋：

- `packages/`：10 個基礎套件，包含 `renderer` 與 `types`
- `packages-plugins/`：19 個內建外掛套件
- `packages-presets/`：3 個樣式預設
- `packages-themes/`：3 個主題套件

## 核心套件

| 套件 | 作用 |
| --- | --- |
| `docx-kit` | 總入口，重匯出公共 API |
| `@docxkit/core` | builder、compiler、樣式系統、外掛契約 |
| `@docxkit/loader` | 動態外掛載入 |
| `@docxkit/pdk` | 外掛開發與測試輔助 |
| `@docxkit/registry` | 外掛探索 |
| `@docxkit/ai` | AI 模板與 prompt 建構 |
| `@docxkit/mcp` | MCP 伺服器 |
| `@docxkit/create-plugin` | 外掛腳手架 |
| `@docxkit/renderer` | 瀏覽器預覽 |
| `@docxkit/types` | 共用型別定義 |
