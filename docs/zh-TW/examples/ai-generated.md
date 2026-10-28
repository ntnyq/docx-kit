# 範例：AI 產生文件

使用 docx-kit 的 AI 範本系統與提示詞建構器，根據結構化參數產生報告的完整範例。

## 完整程式碼

```ts
import {
  createDocx,
  reportTemplate,
  invoiceTemplate,
  resumeTemplate,
  letterTemplate,
  buildPrompt,
  buildFreeformPrompt,
  generateSchemaGuide,
  generateToolDefinitions,
} from 'docx-kit/ai'

// ── 1. Use a built-in template ──
const reportSchema = reportTemplate.generate({
  title: 'Q2 2026 Engineering Report',
  author: 'Engineering Team',
  date: '2026-06-30',
  executiveSummary: [
    'Shipped 14 major features across 3 product lines',
    'Reduced p99 API latency by 35% (320ms → 210ms)',
    'Migrated 100% of services to the new observability stack',
    'Hired 6 senior engineers, 0 regrettable attrition',
  ],
  sections: [
    {
      heading: 'Platform Reliability',
      content: [
        'We achieved 99.97% uptime across all production services, exceeding our 99.95% target.',
        'The migration to the new observability stack is complete; mean time to detection dropped from 12 minutes to 90 seconds.',
      ],
    },
    {
      heading: 'Performance Improvements',
      content: [
        'p99 API latency improved by 35% through a combination of connection pooling, query optimization, and a new caching layer.',
        'We also reduced our CDN egress costs by 22% by moving to a tiered caching strategy.',
      ],
    },
    {
      heading: 'Team & Hiring',
      content: [
        'We welcomed 6 new senior engineers in Q2, bringing the team to 28 total.',
        'Attrition was 0%, well below the industry average of 13%.',
      ],
    },
  ],
  conclusion:
    'Q2 was a strong quarter on all dimensions. We are well-positioned for the H2 roadmap.',
})

const doc = createDocx(reportSchema)
await doc.save('q2-engineering-report.docx')
```

## 範本選項

### 內建範本

docx-kit 提供 4 個 AI 範本：

| 範本 | 使用場景 |
| ----------------- | ----------------------------------------------------------------- |
| `reportTemplate` | 帶執行摘要的商務、工程或財務報告 |
| `invoiceTemplate` | 自動計算小計、稅額與總額的明細發票 |
| `resumeTemplate` | 包含聯繫方式、工作經歷、教育與技能的履歷 |
| `letterTemplate` | 包含發件人、收件人與簽名的正式信函 |

### `invoiceTemplate` 範例

```ts
const invoiceSchema = invoiceTemplate.generate({
  issuerName: 'Acme Corp',
  issuerAddress: '123 Main St, San Francisco, CA 94105',
  clientName: 'Globex Industries',
  clientAddress: '456 Market St, San Francisco, CA 94103',
  invoiceNumber: 'INV-2026-0420',
  issueDate: '2026-06-15',
  dueDate: '2026-07-15',
  taxRate: 0.0875,
  items: [
    {
      description: 'Consulting hours (40 @ $200/hr)',
      quantity: 40,
      unitPrice: 200,
    },
    { description: 'Premium support (annual)', quantity: 1, unitPrice: 2400 },
    { description: 'On-site training (2 days)', quantity: 2, unitPrice: 1500 },
  ],
})

await createDocx(invoiceSchema).save('invoice.docx')
```

### `resumeTemplate` 範例

```ts
const resumeSchema = resumeTemplate.generate({
  name: 'Jane Smith',
  email: 'jane.smith@example.com',
  phone: '+1 555 010 2026',
  summary:
    'Senior software engineer with 10+ years building distributed systems. Passionate about developer experience and operational excellence.',
  experience: [
    {
      role: 'Staff Software Engineer',
      company: 'Acme Corp',
      startDate: '2022-03',
      endDate: 'Present',
      highlights: [
        'Led migration of monolithic platform to 12-service architecture',
        'Reduced p99 API latency by 40% through caching and connection pooling',
        'Mentored 5 engineers; 2 promoted to senior',
      ],
    },
    {
      role: 'Senior Software Engineer',
      company: 'Globex Industries',
      startDate: '2018-06',
      endDate: '2022-02',
      highlights: [
        'Designed and built real-time event-processing pipeline (10K events/sec)',
        'Authored internal best-practices guide adopted org-wide',
      ],
    },
  ],
  education: [
    {
      degree: 'M.S. Computer Science',
      institution: 'Stanford University',
      year: '2018',
    },
    {
      degree: 'B.S. Computer Science',
      institution: 'UC Berkeley',
      year: '2016',
    },
  ],
  skills: [
    'TypeScript',
    'Go',
    'Rust',
    'Kubernetes',
    'PostgreSQL',
    'Kafka',
    'Distributed Systems',
  ],
})

await createDocx(resumeSchema).save('jane-smith-resume.docx')
```

### `letterTemplate` 範例

```ts
const letterSchema = letterTemplate.generate({
  senderName: 'Jane Smith',
  recipientName: 'Hiring Committee',
  subject: 'Application for Senior Engineering Manager',
  body: [
    'I am writing to express my strong interest in the Senior Engineering Manager position. With 10+ years of experience building and leading distributed-systems teams, I am confident I can make a meaningful contribution to your organization.',
    'In my current role at Acme Corp, I lead a team of 12 engineers across two time zones. Under my leadership, we have shipped 14 major features, reduced p99 latency by 35%, and grown the team to 28 with 0% regrettable attrition.',
    'I would welcome the opportunity to discuss how my experience aligns with your needs. Thank you for your consideration.',
  ],
  closing: 'Sincerely',
})

await createDocx(letterSchema).save('cover-letter.docx')
```

## LLM 提示詞建置

使用提示詞建構器為 LLM 產生完整提示詞。

### 基於範本的提示詞

```ts
import { buildPrompt, reportTemplate } from 'docx-kit/ai'

const prompt = buildPrompt(reportTemplate, {
  userContext: 'Focus on infrastructure and reliability metrics',
})
// Returns a complete prompt with:
// 1. Template system prompt
// 2. Template schema reference
// 3. User requirements
// 4. Output format instructions
```

可將 `prompt` 直接發送給 OpenAI、Claude 或其他 LLM，讓模型填寫範本參數。

### 自由格式提示詞（不使用範本）

```ts
import { buildFreeformPrompt } from 'docx-kit/ai'

const prompt = buildFreeformPrompt('business proposal', {
  userContext:
    'For a Series A SaaS startup. Include financial projections and team section.',
})
```

產生通用提示詞，讓 LLM 從零產生任意 docx-kit schema。

## Schema 指南（用於 LLM 上下文）

```ts
import { generateSchemaGuide } from 'docx-kit/ai'

const guide = generateSchemaGuide()
// Returns a long string with:
// - DocxSchema structure overview
// - All block node types with examples
// - All built-in plugins with option schemas
```

將 `guide` 添加到 LLM 提示詞上下文，讓模型瞭解全部可用節點型別。

## 工具定義（OpenAI 函式呼叫）

```ts
import { generateToolDefinitions } from 'docx-kit/ai'

// Default tools: create_document, validate_schema, list_templates
const tools = generateToolDefinitions()

// With plugin listing
const tools2 = generateToolDefinitions({ plugins: ['qrcode', 'echarts'] })
// Adds: list_plugins tool

// With template-specific tool
const tools3 = generateToolDefinitions({ template: 'report' })
// Adds: apply_template_report tool with the template's input schema
```

工具定義遵循 OpenAI 函式呼叫格式，可直接傳給 `openai.chat.completions.create({ tools })`。

## 完整 AI 流程（概念範例）

```ts
import OpenAI from 'openai'
import {
  buildPrompt,
  validateSchema,
  reportTemplate,
  generateToolDefinitions,
} from 'docx-kit/ai'
import { createDocx } from 'docx-kit/node'

const openai = new OpenAI()

// 1. Build the prompt
const prompt = buildPrompt(reportTemplate, {
  userContext: 'Q2 2026 sales numbers — focus on enterprise growth',
})

// 2. Ask the LLM
const response = await openai.chat.completions.create({
  model: 'gpt-4o',
  messages: [
    { role: 'system', content: prompt },
    {
      role: 'user',
      content: 'Generate the Q2 sales report with these numbers: ...',
    },
  ],
  tools: generateToolDefinitions({ template: 'report' }),
})

// 3. Extract & validate the schema
const toolCall = response.choices[0].message.tool_calls[0]
const args = JSON.parse(toolCall.function.arguments)

const validation = validateSchema(args)
if (!validation.valid) {
  console.error('Invalid schema:', validation.errors)
  throw new Error('LLM produced an invalid schema')
}

// 4. Render to .docx
await createDocx(args).save('q2-sales-report.docx')
```

## 相關文件

- [AI 範本與提示詞](/zh-TW/guide/ai-templates) — 完整 API 參考
- [MCP 伺服器端](/zh-TW/guide/mcp-server) — 將 docx-kit 暴露為 MCP 工具
- [JSON DSL](/zh-TW/guide/json-dsl) — 底層 schema 格式
