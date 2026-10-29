# AI 範本與提示詞

docx-kit 提供內建 AI 範本與提示詞建置工具，用於 LLM 驅動的文件產生。

## 安裝

AI 模組透過獨立子路徑匯出：

```ts
import { reportTemplate, buildPrompt } from 'docx-kit/ai'
```

無需額外相依，AI 模組同時支援瀏覽器與 Node.js 環境。

## 內建範本

docx-kit 提供 **4 個內建範本**：

| 範本 | 名稱 | 說明 |
|----------|------|-------------|
| `reportTemplate` | `report` | 包含封面、執行摘要與章節的專業報告 |
| `invoiceTemplate` | `invoice` | 包含費用明細、稅額計算與合計的發票 |
| `resumeTemplate` | `resume` | 包含工作經歷、教育與技能章節的履歷 |
| `letterTemplate` | `letter` | 包含發件人、收件人資訊與簽名區的正式信函 |

### 使用範本

每個範本都有 `generate()` 函式，根據參數產生 `DocxSchema`：

```ts
import { reportTemplate } from 'docx-kit/ai'
import { renderDocx } from 'docx-kit/node'

const schema = reportTemplate.generate({
  title: 'Annual Report',
  author: 'John Doe',
  executiveSummary: ['Key finding 1', 'Key finding 2'],
  sections: [
    { heading: 'Financial Overview', content: ['Revenue increased by 15%.'] },
  ],
  conclusion: 'The outlook for next year is positive.',
})

const doc = await renderDocx(schema)
const buffer = await doc.toBuffer()
```

### 範本參數

#### 報告（`reportTemplate`）

```ts
interface ReportParams {
  title: string               // Document title (required)
  author?: string             // Author name
  date?: string               // Report date
  executiveSummary?: string[] // Summary bullet points
  sections?: {
    heading: string
    content: string[]
  }[]
  conclusion?: string         // Concluding paragraph
}
```

使用核心標題、段落和分頁節點產生可移植的標題頁。

#### 發票（`invoiceTemplate`）

```ts
interface InvoiceParams {
  clientName: string        // Client name (required)
  invoiceNumber: string     // Invoice number (required)
  issuerName: string        // Company name (required)
  items: {
    description: string
    quantity: number
    unitPrice: number
  }[]                       // Line items (required)
  taxRate?: number          // Tax rate (e.g. 0.08 for 8%)
  clientAddress?: string
  issuerAddress?: string
  issueDate?: string
  dueDate?: string
}
```

使用核心表格展示發票中繼資料，自動計算小計、稅額與總額。

#### 履歷（`resumeTemplate`）

```ts
interface ResumeParams {
  name: string              // Full name (required)
  email?: string
  phone?: string
  summary?: string          // Professional summary
  experience?: {
    role?: string
    company?: string
    startDate?: string
    endDate?: string
    highlights?: string[]
  }[]
  education?: {
    degree?: string
    institution?: string
    year?: string
  }[]
  skills?: string[]
}
```

使用核心表格展示聯繫資訊。

#### 信函（`letterTemplate`）

```ts
interface LetterParams {
  senderName: string        // Sender name (required)
  recipientName: string     // Recipient name (required)
  body: string[]            // Letter body paragraphs (required)
  closing?: string          // Closing phrase (default: "Sincerely")
  date?: string
  subject?: string
}
```

使用核心段落節點產生結束語和簽名區域。

## 提示詞建構器

提示詞建構器根據範本或自由格式請求產生可供 LLM 使用的提示詞。

### 基於範本的提示詞

```ts
import { buildPrompt, reportTemplate } from 'docx-kit/ai'

const prompt = buildPrompt(reportTemplate, 'Focus on quarterly revenue trends')
// Returns a complete prompt combining:
// 1. Template system prompt
// 2. Template schema reference (JSON)
// 3. User requirements
// 4. Output format instructions
```

### 自由格式提示詞

適用於不使用範本的文件：

```ts
import { buildFreeformPrompt } from 'docx-kit/ai'

const prompt = buildFreeformPrompt('business proposal', 'Include financial projections')
// Returns a general prompt with:
// 1. Available node types reference
// 2. Available plugins reference
// 3. User requirements
// 4. Output format instructions
```

## Schema 指南

產生全面且適合 LLM 閱讀的 schema 參考：

```ts
import { generateSchemaGuide } from 'docx-kit/ai'

const guide = generateSchemaGuide()
// Returns a multi-line string with:
// - DocxSchema structure overview
// - All block node types with examples
// - All built-in plugins with executable DSL examples
```

## 工具定義

產生兼容 OpenAI 的函式或工具定義，用於 LLM 函式呼叫：

```ts
import { generateToolDefinitions } from 'docx-kit/ai'

// Default tools: create_document, validate_schema, list_templates
const tools = generateToolDefinitions()

// With plugin listing
const tools = generateToolDefinitions({ plugins: ['qrcode', 'echarts'] })
// Adds: list_plugins tool

// With template-specific tool
const tools = generateToolDefinitions({ template: 'report' })
// Adds: apply_template_report tool with the template's input schema
```

工具定義遵循 OpenAI 函式呼叫格式：

```ts
interface ToolDefinition {
  type: 'function'
  function: {
    name: string
    description: string
    parameters: Record<string, unknown>  // JSON Schema format
  }
}
```

## 獨立範本函式

每個範本也可以透過 MCP 模組的 `applyTemplate` 函式應用：

```ts
import { applyTemplate } from 'docx-kit/mcp'

const result = applyTemplate('report', { title: 'My Report' })
if (result) {
  console.log(result.templateName)  // 'report'
  console.log(result.schema)        // DocxSchema object
}
```

## 範本發現

```ts
import { buildTemplateInfoList } from 'docx-kit/mcp'

const templates = buildTemplateInfoList()
// Returns: [
//   { name: 'report', description: '...', systemPrompt: '...', schema: {...} },
//   { name: 'invoice', ... },
//   { name: 'resume', ... },
//   { name: 'letter', ... },
// ]
```


## 範本輸入驗證

`generateToolDefinitions()` 輸出標準 JSON Schema，並轉換巢狀必填欄位。自訂整合可用 `toTemplateJsonSchema()` 轉換範本欄位定義，以 `validateTemplateData()` 驗證未知輸入。`applyTemplate()` 會先驗證再產生，參數無效時拋出包含 JSON Pointer 欄位路徑的 `TypeError`。

```ts
import { reportTemplate, toTemplateJsonSchema, validateTemplateData } from 'docx-kit/ai'

const inputSchema = toTemplateJsonSchema(reportTemplate.schema)
const errors = validateTemplateData(reportTemplate.schema, {})
// [{ path: '/title', message: 'Required field is missing' }]
```
