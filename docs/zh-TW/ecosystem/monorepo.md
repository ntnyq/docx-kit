# Monorepo 結構

docx-kit 以 **pnpm workspace monorepo** 維護。文件涵蓋目前工作區中的核心套件、22 個外掛套件、3 個預設與 3 個主題。

## 目錄結構

```text
docx-kit/
├── packages/
│   ├── ai/
│   ├── core/
│   ├── create-plugin/
│   ├── docx-kit/
│   ├── loader/
│   ├── markdown/
│   ├── mcp/
│   ├── pdk/
│   ├── registry/
│   ├── renderer/
│   └── types/
├── packages-plugins/
│   ├── badge/
│   ├── barcode/
│   ├── callout/
│   ├── changelog/
│   ├── checklist/
│   ├── code-block/
│   ├── cover-page/
│   ├── data-table/
│   ├── divider/
│   ├── echarts/
│   ├── figure/
│   ├── invoice/
│   ├── letterhead/
│   ├── meeting-minutes/
│   ├── metric-card/
│   ├── page-number/
│   ├── property-table/
│   ├── qrcode/
│   ├── signature-block/
│   ├── timeline/
│   ├── toc/
│   └── watermark/
├── packages-presets/
│   ├── academic/
│   ├── classic/
│   └── modern/
├── packages-themes/
│   ├── minimal/
│   ├── ocean/
│   └── warm/
└── docs/
```

## 工作區包分佈

| 範圍 | 數量 | 說明 |
| ------------------- | ----- | ------------------------------------------------------------ |
| `packages/` | 11 | 執行階段、工具、AI、呈現器、Markdown 與共享型別 |
| `packages-plugins/` | 22 | 每個內建外掛都有獨立的套件、測試與 README |
| `packages-presets/` | 3 | 常見文件型別的樣式預設 |
| `packages-themes/` | 3 | 基於令牌的配色主題 |

## `packages/` 中的內容

| 套件 | 職責 |
| ------------------------ | --------------------------------------------------- |
| `docx-kit` | 重新匯出公共 API 的整合套件 |
| `@docxkit/core` | 建構器、編譯器、樣式系統與外掛契約 |
| `@docxkit/loader` | 從多種來源動態載入外掛 |
| `@docxkit/pdk` | 外掛開發與測試輔助工具 |
| `@docxkit/registry` | 基於 npm 的外掛發現 |
| `@docxkit/ai` | 範本、schema 指引與提示詞建置 |
| `@docxkit/mcp` | 面向 AI Agent 整合的 MCP 伺服器端 |
| `@docxkit/create-plugin` | 外掛鷹架 CLI |
| `@docxkit/renderer` | 瀏覽器端 `.docx` 預覽 |
| `@docxkit/types` | 共享的文件、樣式、外掛與工具型別 |

[套件總覽](/zh-TW/ecosystem/packages)提供各個包的詳細使用建議。

## 外掛封裝規則

內建外掛遵循統一模式：

- 套件名稱：`@docxkit/plugin-{name}`
- 目錄：`packages-plugins/{name}`
- 將 `@docxkit/core` 作為同儕相依套件
- 使用 `tsdown` 建置，設定 `platform: 'neutral'`
- 每個外掛都有獨立 README 與測試

外掛可能有不同的同儕相依套件和執行階段限制，因此保持獨立套件很重要。

## 預設與主題

- 預設提供 `classic`、`modern`、`academic` 等明確的基本樣式。
- 主題提供 `minimal`、`ocean`、`warm` 等可重用令牌。
- 整合套件會重新匯出它們，除非需要按套件控制相依，否則可統一使用 `docx-kit`。

## 貢獻流程

使用 `.node-version` 指定的 Node.js LTS，以及根目錄
`packageManager` 欄位固定的 pnpm 版本。首次安裝前啟用 Corepack。
完整的審查與兼容性策略見[倉庫
貢獻
指南](https://github.com/ntnyq/docx-kit/blob/main/CONTRIBUTING.md)。

在倉庫根目錄執行：

```sh
corepack enable
corepack install
pnpm install --frozen-lockfile
pnpm run release:verify
```

常用的專項檢查命令：

```sh
pnpm exec vitest run packages/core/tests/compiler/ooxml-contracts.test.ts
pnpm exec vitest run tests/e2e/plugin-lifecycle.test.ts
pnpm run visual:check
pnpm run bench:node
pnpm run bench:browser
```

視覺和性能檢查的臨時輸出寫入已忽略的 `tmp/` 與
`output/` 目錄。只有確認變更符合預期後，
才提交視覺或基準測試基線。

## 文件維護要求

工作區變化時，同步更新以下三個位置：

- 首頁功能摘要
- 外掛與套件總覽
- 所有支援語言的導航

這樣文件才能與實際 monorepo 保持一致。
