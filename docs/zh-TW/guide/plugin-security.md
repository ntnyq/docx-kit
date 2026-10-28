# 外掛安全指南

docx-kit 的外掛系統支援從 npm 套件、URL 與本地檔案載入外部程式碼。本指南介紹其信任模型、權限鉤子，以及安全使用和編寫外掛的實踐。

## 信任級別

按來源與驗證方式，外掛分為四個信任級別：

| 級別 | 來源 | 驗證方式 | 風險 |
|--------------|-------------------------------------------|-------------------------------------------|---------|
| **內建** | 包含在 docx-kit 原始碼中 | 每次發佈均經過測試和審查 | 極低 |
| **已驗證** | 由 docx-kit 團隊發佈 | 驗證清單與版本 | 低 |
| **社區** | 第三方發佈到 npm | 驗證清單與關鍵字標籤 | 中 |
| **不受信任** | 任意 URL 或本地路徑 | 驗證清單，但不信任發佈者 | 高 |

內建外掛（callout、watermark、echarts、qrcode 等）隨 docx-kit 發佈週期一起編譯和測試。已驗證外掛遵循相同審查流程，但位於獨立套件。社區、URL 和本地外掛接受清單級驗證。URL 與本地外掛風險最高，因為清單有效並不代表可執行程式碼可信。

## 權限鉤子

`PluginLoader` 提供兩個安全鉤子，讓使用者控制載入和執行哪些外掛：

### `allowLoad(source)`

在載入開始**之前**呼叫，接收 `PluginSource` 描述符。回傳 `false` 可完全阻止載入。

```ts
import { createPluginLoader } from 'docx-kit/loader'

const loader = createPluginLoader({
  security: {
    // Only allow npm packages from trusted registries
    allowLoad(source) {
      if (source.type === 'url') {
        const domain = new URL(source.url).origin
        return domain === 'https://trusted.example.com'
      }
      return true // allow npm and inline sources
    },
  },
})
```

適合設置域名白名單、限制 npm 套件名，或在禁止網絡請求的環境中阻止 URL 載入。

### `allowExecute(manifest, source)`

在解析清單**之後**、執行外掛程式碼**之前**呼叫，接收解析後的 `PluginManifest` 與原始 `PluginSource`。回傳 `false` 可阻止執行。

```ts
const loader = createPluginLoader({
  security: {
    // Only execute plugins with compatible version ranges
    allowExecute(manifest, source) {
      // Block plugins that require exact versions (no semver ranges)
      if (!manifest.docxKit.startsWith('^') && manifest.docxKit !== '*') {
        console.warn(`Blocked plugin with pinned version: ${manifest.plugin.name}`)
        return false
      }
      return true
    },
  },
})
```

可在程式碼執行前檢查清單中繼資料（作者、版本範圍、相依），以執行版本策略、拒絕無人維護的套件或阻止相依清單可疑的外掛。

## 清單驗證

預設情況下，從 npm、URL 或本地來源載入的每個外掛都經過 `validateManifest()`，確保：

- 必填欄位（`name`、`version`、`docxKit`、`main`、`plugin.name`）存在且型別正確
- `version` 為有效 semver（如 `1.0.0`）
- `docxKit` 為有效 semver 範圍（如 `^0.2.0`、`~0.3.0`、`*`）
- `main` 和可選 `types` 為以 `./` 開頭的安全套件內相對路徑
- 可選相依映射與匯出映射只包含字串值

載入可信或測試外掛時，可以關閉清單驗證：

```ts
const loader = createPluginLoader({
  validateManifest: false, // skip manifest checks
})
```

**警告：** 關閉驗證僅適用於本地開發或完全由你控制的外掛。不要為 URL 或社區外掛關閉驗證。

## 版本兼容性

載入器會檢查外掛的 `docxKit` 範圍是否匹配已安裝的 docx-kit 版本：

| 範圍 | 含義 | 範例說明 |
|-------------|----------------------------------------------|------------------------|
| `*` | 任意版本，始終兼容 | 通配符 |
| `^0.2.0` | >= 0.2.0 且 < 0.3.0 | 脫字符範圍（0.x 版本保持次版本兼容） |
| `~0.2.0` | >= 0.2.0 且 < 0.3.0 | 波浪號範圍，允許補丁版本更新 |

若外掛聲明 `^0.2.0`，但已安裝版本為 `1.0.0`，載入器會拋出 `PLUGIN_VERSION_MISMATCH`，避免宿主庫升級後外掛靜默失效。

## 使用者最佳實踐

1. **始終驗證清單。** 對於未完全控制的外掛來源，保留預設的 `validateManifest: true`。

2. **限制 URL 來源。** 使用 `allowLoad()` 設置域名白名單，生產環境不要從任意 URL 載入外掛。

3. **檢查版本範圍。** 使用 `allowExecute()` 拒絕過於精確或使用通配符的範圍，避免掩蓋兼容性問題。

4. **儘量從 npm 載入。** npm 套件至少經過登錄服務驗證，並可透過 `keyword:docx-kit-plugin` 發現。使用登錄服務搜尋外掛：

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const plugins = await registry.search('chart')
   ```

5. **使用 `loadAll()` 批量載入。** 它收整合功結果，對失敗發出警告而不拋出異常，適合啟動流程：

   ```ts
   const results = await loader.loadAll([
     { plugin: calloutPlugin(), type: 'inline' },
     { package: 'docx-kit-plugin-chart', type: 'npm' },
   ])
   ```

## 作者最佳實踐

1. **始終提供清單。** 在套件根目錄建立 `docx-kit.plugin.json`，以便載入器驗證並檢查兼容性。

2. **為 `docxKit` 聲明脫字符範圍。** 使用 `^0.2.0`（或目前主版本、次版本），避免 `*` 或固定版本，以表達兼容意圖並保護使用者。

3. **添加 npm 關鍵字。** 在 `package.json` 的 keywords 中添加 `docx-kit-plugin`，讓登錄服務搜尋能夠發現外掛。

4. **聲明同儕相依套件。** 將 `docx-kit` 放入 peerDependency，而非普通相依，避免使用者項目出現版本衝突。

5. **使用 PDK 編寫測試。** 透過 `docx-kit/pdk` 的 `renderPlugin()` 和 `assertRendersParagraph()` 獨立測試外掛：

   ```ts
   import { renderPlugin, assertRendersParagraph } from 'docx-kit/pdk'
   import { myPlugin } from './src'

   const result = await renderPlugin(myPlugin(), { text: 'Hello' })
   assertRendersParagraph(result, 'Hello')
   ```

## 沙箱（未來計劃）

目前外掛程式碼與宿主應用在同一上下文執行。未來版本可能透過以下方式引入沙箱：

- 瀏覽器環境的 **Web Workers**
- Node.js 的 **VM 模組**（`node:vm`）
- 結合**內容安全策略**，透過 iframe 隔離

沙箱支援推出前，應主要相依權限鉤子和清單驗證進行防護。
