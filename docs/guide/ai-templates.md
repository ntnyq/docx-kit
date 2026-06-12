# AI Templates & Prompts

docx-kit provides built-in AI templates and prompt-building utilities for LLM-powered document generation.

## Installation

The AI module is available as a separate subpath export:

```ts
import { reportTemplate, buildPrompt } from 'docx-kit/ai'
```

No additional dependencies are required — the AI module works in both browser and Node.js environments.

## Built-in Templates

docx-kit ships with **4 built-in templates**:

| Template | Name | Description |
|----------|------|-------------|
| `reportTemplate` | `report` | Professional report with cover page, executive summary, and sections |
| `invoiceTemplate` | `invoice` | Invoice with itemized charges, tax calculation, and totals |
| `resumeTemplate` | `resume` | Resume/CV with experience, education, and skills sections |
| `letterTemplate` | `letter` | Formal letter with sender/recipient info and signature block |

### Using a Template

Each template has a `generate()` function that produces a `DocxSchema` from parameters:

```ts
import { reportTemplate, createDocx } from 'docx-kit'
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

const doc = createDocx(schema)
const buffer = await renderDocx(doc)
```

### Template Parameters

#### Report (`reportTemplate`)

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

Uses the `coverPage` plugin for the title page.

#### Invoice (`invoiceTemplate`)

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

Uses the `propertyTable` plugin for invoice metadata and calculates subtotal/tax/total automatically.

#### Resume (`resumeTemplate`)

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

Uses the `propertyTable` plugin for contact info.

#### Letter (`letterTemplate`)

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

Uses the `signatureBlock` plugin for the closing/signature area.

## Prompt Builder

The prompt builder creates LLM-ready prompts from templates or freeform requests.

### Template-based Prompt

```ts
import { buildPrompt, reportTemplate } from 'docx-kit/ai'

const prompt = buildPrompt(reportTemplate, 'Focus on quarterly revenue trends')
// Returns a complete prompt combining:
// 1. Template system prompt
// 2. Template schema reference (JSON)
// 3. User requirements
// 4. Output format instructions
```

### Freeform Prompt

For documents without a template:

```ts
import { buildFreeformPrompt } from 'docx-kit/ai'

const prompt = buildFreeformPrompt('business proposal', 'Include financial projections')
// Returns a general prompt with:
// 1. Available node types reference
// 2. Available plugins reference
// 3. User requirements
// 4. Output format instructions
```

## Schema Guide

Generate a comprehensive LLM-friendly schema reference:

```ts
import { generateSchemaGuide } from 'docx-kit/ai'

const guide = generateSchemaGuide()
// Returns a multi-line string with:
// - DocxSchema structure overview
// - All block node types with examples
// - All built-in plugins with option schemas
```

## Tool Definitions

Generate OpenAI-compatible function/tool definitions for LLM function calling:

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

The tool definitions follow the OpenAI function calling format:

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

## Standalone Template Functions

Each template can also be applied via the `applyTemplate` function from the MCP module:

```ts
import { applyTemplate } from 'docx-kit/mcp'

const result = applyTemplate('report', { title: 'My Report' })
if (result) {
  console.log(result.templateName)  // 'report'
  console.log(result.schema)        // DocxSchema object
}
```

## Template Discovery

```ts
import { buildTemplateInfoList } from 'docx-kit/ai'

const templates = buildTemplateInfoList()
// Returns: [
//   { name: 'report', description: '...', systemPrompt: '...', schema: {...} },
//   { name: 'invoice', ... },
//   { name: 'resume', ... },
//   { name: 'letter', ... },
// ]
```