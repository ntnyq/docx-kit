# MCP 伺服器端

docx-kit 提供模型上下文協議（MCP）伺服器端，讓 Claude、ChatGPT 和其他兼容 MCP 的 AI Agent 能夠建立、驗證和探索 docx-kit 文件。

## 安裝

MCP 伺服器端需要額外的可選同儕相依套件：

```bash
pnpm add @modelcontextprotocol/sdk zod
```

## 快速開始

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'

const server = await createDocxKitServer()
const transport = new StdioServerTransport()
await server.connect(transport)
```

## 可用工具

MCP 伺服器端提供 **6 個工具**：

### `create_document`

根據 docx-kit JSON schema 建立新的 .docx 文件。

**參數：**
- `outputPath`（string）：輸出 .docx 檔案的路徑
- `schema`（object）：docx-kit DocxSchema JSON 物件

**回傳：** 產生檔案的絕對 `filePath` 與位元組數 `size`。

21 個兼容 Node.js 的內建外掛會自動註冊。ECharts
需要瀏覽器，因此不會出現在 `list_plugins` 中。外部外掛來源
仍需要顯式設定 `pluginLoader`。

輸出檔案限制在 `outputDirectory` 內，預設使用工作目錄。
該目錄及其父目錄應由伺服器端所有者控制；請求執行期間，
不要允許不受信任的本地進程替換目錄。

### `validate_schema`

驗證 docx-kit JSON schema 是否正確。

**參數：**
- `schema`（object）：待驗證的 DocxSchema JSON 物件

**回傳：** 驗證狀態；無效時附帶詳細錯誤資訊。

檢查內容：
- 必填的 `content` 陣列
- 有效節點型別（heading、paragraph、table 等）
- 各節點型別的必填欄位
- 外掛名稱引用

### `list_plugins`

列出 Node.js 伺服器端中可用的內建 docx-kit 外掛。

**參數：**
- `filter`（string，可選）：按外掛名稱模式篩選

**回傳：** 包含名稱和說明的外掛資訊物件陣列。

### `get_plugin_help`

獲取指定外掛的使用幫助與範例。

**參數：**
- `pluginName`（string）：外掛名稱

**回傳：** 包含說明、用法範例與選項 schema 的外掛幫助資訊。

### `list_templates`

列出全部可用的 docx-kit AI 範本。

**回傳：** 包含名稱、說明、系統提示詞與輸入 schema 的範本資訊物件陣列。

### `apply_template`

根據 AI 範本和參數資料產生文件 schema。

**參數：**
- `template`（string）：範本名稱（`report`、`invoice`、`resume`、`letter`）
- `data`（object）：範本參數值

**回傳：** 產生的 DocxSchema 物件；找不到範本時回傳錯誤。

## 資源

### `docx-kit://schema`

描述完整 DocxSchema 結構的 JSON Schema 資源，包括：
- 文件級欄位（content、styles、page、plugins）
- 全部區塊級節點型別及其必填和可選欄位
- 外掛選項 schema

MCP 用戶端可讀取該資源，瞭解預期的文件格式。

## 編程呼叫（無需 MCP SDK）

所有工具邏輯函式均作為獨立 API 匯出，無需 MCP SDK 即可使用：

```ts
import {
  validateSchema,
  applyTemplate,
  buildPluginInfoList,
  buildPluginHelp,
  buildTemplateInfoList,
} from 'docx-kit/mcp'

// Validate a schema
const result = validateSchema({
  content: [{ type: 'heading', level: 1, text: 'Title' }],
})
console.log(result.valid)   // true
console.log(result.errors)  // []

// Apply a template
const doc = applyTemplate('report', { title: 'Annual Report' })
console.log(doc?.templateName)  // 'report'
console.log(doc?.schema)        // DocxSchema object

// Build plugin info
const plugins = [{ name: 'qrcode', render: () => 'qr' }]
const info = buildPluginInfoList(plugins, 'code')  // Filtered by name
console.log(info)  // [{ name: 'qrcode', description: '...' }]

// Build plugin help
const help = buildPluginHelp({ name: 'qrcode', render: () => 'qr' })
console.log(help.usageExample)  // '{ type: "plugin", name: "qrcode", ... }'

// List templates
const templates = buildTemplateInfoList()
console.log(templates.map(t => t.name))  // ['invoice', 'letter', 'report', 'resume']
```

## 工具定義

還匯出了 MCP 工具定義，供自訂 MCP 伺服器端實現使用：

```ts
import { TOOL_DEFINITIONS, RESOURCE_DEFINITIONS } from 'docx-kit/mcp'

for (const def of TOOL_DEFINITIONS) {
  console.log(def.name, def.description)
}
// validate_schema - Validate a docx-kit JSON schema...
// create_document - Create a new .docx document...
// etc.
```

## HTTP 傳輸

透過 HTTP 訪問時，使用 StreamableHTTPServerTransport：

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'

const server = await createDocxKitServer()
const transport = new StreamableHTTPServerTransport({ port: 3000 })
await server.connect(transport)
```

## 架構

MCP 伺服器端模組支援兩種使用方式：

- **協議模式**：透過 `createDocxKitServer()` 提供完整 MCP 伺服器端，需要 `@modelcontextprotocol/sdk` 和 `zod`
- **編程模式**：透過獨立匯出函式呼叫，無需 MCP SDK

`src/mcp.ts` 入口重新匯出所有內容；僅需要特定函式時，也可以直接匯入 `src/mcp-server/tools/` 中的獨立工具模組。

## 錯誤處理

`validateSchema` 函式回傳結構化錯誤資訊：

```ts
interface ValidationError {
  message: string  // Error description
  path: string     // JSON path (e.g. '/content/0/type')
}

interface ValidateSchemaOutput {
  valid: boolean
  errors: ValidationError[]
}
```

常見驗證錯誤：

| 路徑 | 消息 |
|------|---------|
| `/content` | 缺少必填欄位 "content" |
| `/content` | 欄位 "content" 必須是陣列 |
| `/content/N/type` | 索引 N 處的節點缺少必填欄位 "type" |
| `/content/N/type` | 索引 N 處的節點型別 "X" 無效 |
| `/content/N/text` | 標題節點缺少必填欄位 "text" |
| `/content/N/level` | 標題節點缺少必填欄位 "level" |
| `/content/N/name` | 外掛節點缺少必填欄位 "name" |
