# MCP Server

docx-kit provides a Model Context Protocol (MCP) server for AI agent integration. This enables LLM-powered tools like Claude, ChatGPT, and other MCP-compatible agents to create, validate, and explore docx-kit documents.

## Installation

The MCP server requires additional optional peer dependencies:

```bash
pnpm add @modelcontextprotocol/sdk zod
```

## Quick Start

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'

const server = await createDocxKitServer()
const transport = new StdioServerTransport()
await server.connect(transport)
```

## Available Tools

The MCP server provides **6 tools**:

### `create_document`

Create a new .docx document from a docx-kit JSON schema.

**Parameters:**
- `outputPath` (string): File path for the output .docx file
- `schema` (object): A docx-kit DocxSchema JSON object

**Returns:** The created file's absolute `filePath` and byte `size`.

The 18 Node-compatible built-in plugins are registered automatically. ECharts
requires a browser and is not included in `list_plugins`. External plugin sources
still require an explicitly configured `pluginLoader`.

Files are restricted to `outputDirectory` (the working directory by default).
Keep this directory and its ancestors under the server owner's control; do not
allow untrusted local processes to replace directories while requests run.

### `validate_schema`

Validate a docx-kit JSON schema for correctness.

**Parameters:**
- `schema` (object): The DocxSchema JSON object to validate

**Returns:** Validation status with detailed error information if invalid.

Checks for:
- Required `content` array
- Valid node types (heading, paragraph, table, etc.)
- Required fields per node type
- Plugin name references

### `list_plugins`

List the built-in docx-kit plugins available in the Node.js server.

**Parameters:**
- `filter` (string, optional): Filter by plugin name pattern

**Returns:** Array of plugin info objects with name and description.

### `get_plugin_help`

Get usage help and examples for a specific plugin.

**Parameters:**
- `pluginName` (string): The plugin name

**Returns:** Plugin help info with description, usage example, and option schema.

### `list_templates`

List all available docx-kit AI templates.

**Returns:** Array of template info objects with name, description, system prompt, and input schema.

### `apply_template`

Generate a document schema from an AI template + parameter data.

**Parameters:**
- `template` (string): Template name (`report`, `invoice`, `resume`, `letter`)
- `data` (object): Template parameter values

**Returns:** Generated DocxSchema object, or error if template not found.

## Resources

### `docx-kit://schema`

A JSON Schema resource describing the complete DocxSchema structure, including:
- Document-level fields (content, styles, page, plugins)
- All block node types and their required/optional fields
- Plugin option schemas

MCP clients can read this resource to understand the expected document format.

## Programmatic Use (without MCP SDK)

All tool logic functions are available as standalone exports, usable without the MCP SDK:

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

## Tool Definitions

MCP tool definitions are also exported for custom MCP server implementations:

```ts
import { TOOL_DEFINITIONS, RESOURCE_DEFINITIONS } from 'docx-kit/mcp'

for (const def of TOOL_DEFINITIONS) {
  console.log(def.name, def.description)
}
// validate_schema - Validate a docx-kit JSON schema...
// create_document - Create a new .docx document...
// etc.
```

## HTTP Transport

For HTTP-based access, use the StreamableHTTPServerTransport:

```ts
import { createDocxKitServer } from 'docx-kit/mcp'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'

const server = await createDocxKitServer()
const transport = new StreamableHTTPServerTransport({ port: 3000 })
await server.connect(transport)
```

## Architecture

The MCP server module is structured for dual use:

- **Protocol mode**: Full MCP server via `createDocxKitServer()` — requires `@modelcontextprotocol/sdk` and `zod`
- **Programmatic mode**: Standalone functions via individual exports — no MCP SDK required

The `src/mcp.ts` entry point re-exports everything, while individual tool modules in `src/mcp-server/tools/` can be imported directly if you only need specific functions.

## Error Handling

The `validateSchema` function returns structured error information:

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

Common validation errors:

| Path | Message |
|------|---------|
| `/content` | Required field "content" is missing |
| `/content` | Field "content" must be an array |
| `/content/N/type` | Node at index N missing required "type" field |
| `/content/N/type` | Invalid node type "X" at index N |
| `/content/N/text` | Heading node missing required "text" field |
| `/content/N/level` | Heading node missing required "level" field |
| `/content/N/name` | Plugin node missing required "name" field |
