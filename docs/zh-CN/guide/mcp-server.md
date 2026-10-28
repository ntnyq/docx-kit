# MCP 服务端

docx-kit 提供模型上下文协议（MCP）服务端，让 Claude、ChatGPT 和其他兼容 MCP 的 AI Agent 能够创建、校验和探索 docx-kit 文档。

## 安装

MCP 服务端需要额外的可选对等依赖：

```bash
pnpm add @modelcontextprotocol/sdk zod
```

## 快速开始

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'

const server = await createDocxKitServer()
const transport = new StdioServerTransport()
await server.connect(transport)
```

## 可用工具

MCP 服务端提供 **6 个工具**：

### `create_document`

根据 docx-kit JSON schema 创建新的 .docx 文档。

**参数：**
- `outputPath`（string）：输出 .docx 文件的路径
- `schema`（object）：docx-kit DocxSchema JSON 对象

**返回：** 生成文件的绝对 `filePath` 与字节数 `size`。

21 个兼容 Node.js 的内置插件会自动注册。ECharts
需要浏览器，因此不会出现在 `list_plugins` 中。外部插件来源
仍需要显式配置 `pluginLoader`。

输出文件限制在 `outputDirectory` 内，默认使用工作目录。
该目录及其父目录应由服务端所有者控制；请求运行期间，
不要允许不受信任的本地进程替换目录。

### `validate_schema`

校验 docx-kit JSON schema 是否正确。

**参数：**
- `schema`（object）：待校验的 DocxSchema JSON 对象

**返回：** 校验状态；无效时附带详细错误信息。

检查内容：
- 必填的 `content` 数组
- 有效节点类型（heading、paragraph、table 等）
- 各节点类型的必填字段
- 插件名称引用

### `list_plugins`

列出 Node.js 服务端中可用的内置 docx-kit 插件。

**参数：**
- `filter`（string，可选）：按插件名称模式筛选

**返回：** 包含名称和说明的插件信息对象数组。

### `get_plugin_help`

获取指定插件的使用帮助与示例。

**参数：**
- `pluginName`（string）：插件名称

**返回：** 包含说明、用法示例与选项 schema 的插件帮助信息。

### `list_templates`

列出全部可用的 docx-kit AI 模板。

**返回：** 包含名称、说明、系统提示词与输入 schema 的模板信息对象数组。

### `apply_template`

根据 AI 模板和参数数据生成文档 schema。

**参数：**
- `template`（string）：模板名称（`report`、`invoice`、`resume`、`letter`）
- `data`（object）：模板参数值

**返回：** 生成的 DocxSchema 对象；找不到模板时返回错误。

## 资源

### `docx-kit://schema`

描述完整 DocxSchema 结构的 JSON Schema 资源，包括：
- 文档级字段（content、styles、page、plugins）
- 全部块级节点类型及其必填和可选字段
- 插件选项 schema

MCP 客户端可读取该资源，了解预期的文档格式。

## 编程调用（无需 MCP SDK）

所有工具逻辑函数均作为独立 API 导出，无需 MCP SDK 即可使用：

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

## 工具定义

还导出了 MCP 工具定义，供自定义 MCP 服务端实现使用：

```ts
import { TOOL_DEFINITIONS, RESOURCE_DEFINITIONS } from 'docx-kit/mcp'

for (const def of TOOL_DEFINITIONS) {
  console.log(def.name, def.description)
}
// validate_schema - Validate a docx-kit JSON schema...
// create_document - Create a new .docx document...
// etc.
```

## HTTP 传输

通过 HTTP 访问时，使用 StreamableHTTPServerTransport：

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'

const server = await createDocxKitServer()
const transport = new StreamableHTTPServerTransport({ port: 3000 })
await server.connect(transport)
```

## 架构

MCP 服务端模块支持两种使用方式：

- **协议模式**：通过 `createDocxKitServer()` 提供完整 MCP 服务端，需要 `@modelcontextprotocol/sdk` 和 `zod`
- **编程模式**：通过独立导出函数调用，无需 MCP SDK

`src/mcp.ts` 入口重新导出所有内容；仅需要特定函数时，也可以直接导入 `src/mcp-server/tools/` 中的独立工具模块。

## 错误处理

`validateSchema` 函数返回结构化错误信息：

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

常见校验错误：

| 路径 | 消息 |
|------|---------|
| `/content` | 缺少必填字段 "content" |
| `/content` | 字段 "content" 必须是数组 |
| `/content/N/type` | 索引 N 处的节点缺少必填字段 "type" |
| `/content/N/type` | 索引 N 处的节点类型 "X" 无效 |
| `/content/N/text` | 标题节点缺少必填字段 "text" |
| `/content/N/level` | 标题节点缺少必填字段 "level" |
| `/content/N/name` | 插件节点缺少必填字段 "name" |
