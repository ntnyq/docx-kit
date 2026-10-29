# AI 模板与提示词

docx-kit 提供内置 AI 模板与提示词构建工具，用于 LLM 驱动的文档生成。

## 安装

AI 模块通过独立子路径导出：

```ts
import { reportTemplate, buildPrompt } from 'docx-kit/ai'
```

无需额外依赖，AI 模块同时支持浏览器与 Node.js 环境。

## 内置模板

docx-kit 提供 **4 个内置模板**：

| 模板 | 名称 | 说明 |
|----------|------|-------------|
| `reportTemplate` | `report` | 包含封面、执行摘要与章节的专业报告 |
| `invoiceTemplate` | `invoice` | 包含费用明细、税额计算与合计的发票 |
| `resumeTemplate` | `resume` | 包含工作经历、教育与技能章节的简历 |
| `letterTemplate` | `letter` | 包含发件人、收件人信息与签名区的正式信函 |

### 使用模板

每个模板都有 `generate()` 函数，根据参数生成 `DocxSchema`：

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

### 模板参数

#### 报告（`reportTemplate`）

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

使用核心标题、段落和分页节点生成可移植的标题页。

#### 发票（`invoiceTemplate`）

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

使用核心表格展示发票元数据，自动计算小计、税额与总额。

#### 简历（`resumeTemplate`）

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

使用核心表格展示联系信息。

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

使用核心段落节点生成结束语和签名区域。

## 提示词构建器

提示词构建器根据模板或自由格式请求生成可供 LLM 使用的提示词。

### 基于模板的提示词

```ts
import { buildPrompt, reportTemplate } from 'docx-kit/ai'

const prompt = buildPrompt(reportTemplate, 'Focus on quarterly revenue trends')
// Returns a complete prompt combining:
// 1. Template system prompt
// 2. Template schema reference (JSON)
// 3. User requirements
// 4. Output format instructions
```

### 自由格式提示词

适用于不使用模板的文档：

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

生成全面且适合 LLM 阅读的 schema 参考：

```ts
import { generateSchemaGuide } from 'docx-kit/ai'

const guide = generateSchemaGuide()
// Returns a multi-line string with:
// - DocxSchema structure overview
// - All block node types with examples
// - All built-in plugins with executable DSL examples
```

## 工具定义

生成兼容 OpenAI 的函数或工具定义，用于 LLM 函数调用：

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

工具定义遵循 OpenAI 函数调用格式：

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

## 独立模板函数

每个模板也可以通过 MCP 模块的 `applyTemplate` 函数应用：

```ts
import { applyTemplate } from 'docx-kit/mcp'

const result = applyTemplate('report', { title: 'My Report' })
if (result) {
  console.log(result.templateName)  // 'report'
  console.log(result.schema)        // DocxSchema object
}
```

## 模板发现

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


## 模板输入校验

`generateToolDefinitions()` 输出标准 JSON Schema，并转换嵌套必填字段。自定义集成可用 `toTemplateJsonSchema()` 转换模板字段定义，用 `validateTemplateData()` 校验未知输入。`applyTemplate()` 会先校验再生成，参数无效时抛出包含 JSON Pointer 字段路径的 `TypeError`。

```ts
import { reportTemplate, toTemplateJsonSchema, validateTemplateData } from 'docx-kit/ai'

const inputSchema = toTemplateJsonSchema(reportTemplate.schema)
const errors = validateTemplateData(reportTemplate.schema, {})
// [{ path: '/title', message: 'Required field is missing' }]
```
