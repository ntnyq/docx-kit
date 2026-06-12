/**
 * Generate LLM-friendly schema documentation and
 * OpenAI-compatible tool definitions.
 *
 * @module ai/schemaGuide
 */

import { BUILTIN_TEMPLATES } from './templates'
import type { AiTemplate, GenerateToolDefinitionsOptions } from './types'

/**
 * Generate a human-readable schema guide for LLMs.
 *
 * Produces a comprehensive reference document describing
 * all available node types, plugins, and their options.
 *
 * @returns Schema guide text
 */
export function generateSchemaGuide(): string {
  const parts = [
    '# docx-kit Schema Reference for LLMs',
    '',
    '## Document Schema Structure',
    'A docx-kit document is defined by a `DocxSchema` object:',
    '',
    '```json',
    '{',
    '  "content": [ /* array of BlockNode objects */ ],',
    '  "styles": { /* optional: className → style rule map */ },',
    '  "page": { /* optional: page size, margins, orientation */ },',
    '  "plugins": [ /* optional: plugin sources to load */ ]',
    '}',
    '```',
    '',
    '## Block Node Types',
    '',
    '### heading',
    '```json',
    '{ "type": "heading", "level": 1, "text": "Title", "className": "h1", "style": { "fontSize": 24 } }',
    '```',
    '- `level`: 1–6 (required)',
    '- `text`: heading text (required)',
    '- `className`, `style`: optional overrides',
    '',
    '### paragraph',
    '```json',
    '{ "type": "paragraph", "text": "Body text", "className": "p" }',
    '```',
    '- `text`: plain text content',
    '- `children`: array of inline nodes (TextNode, ImageNode, HyperlinkNode)',
    '- Either `text` or `children` must be present',
    '',
    '### table',
    '```json',
    '{',
    '  "type": "table",',
    '  "columns": [{ "key": "name", "title": "Name" }, { "key": "value", "title": "Value" }],',
    '  "data": [{ "name": "Item", "value": "$10" }],',
    '  "header": true,',
    '  "bordered": true',
    '}',
    '```',
    '',
    '### bulletList',
    '```json',
    '{ "type": "bulletList", "items": ["Item 1", "Item 2"] }',
    '```',
    '',
    '### numberedList',
    '```json',
    '{ "type": "numberedList", "items": ["Step 1", "Step 2"], "start": 1 }',
    '```',
    '',
    '### image',
    '```json',
    '{ "type": "image", "data": "<base64 or URL>", "width": 200, "height": 100, "alt": "Logo" }',
    '```',
    '',
    '### pageBreak',
    '```json',
    '{ "type": "pageBreak" }',
    '```',
    '',
    '### hyperlink',
    '```json',
    '{ "type": "hyperlink", "url": "https://example.com", "children": ["Click here"] }',
    '```',
    '',
    '### plugin (dynamic)',
    '```json',
    '{ "type": "plugin", "name": "pluginName", "options": { /* plugin-specific */ } }',
    '```',
    '',
    '## Built-in Plugins',
  ]

  const pluginDescriptions: [string, string, string][] = [
    [
      'callout',
      'Highlighted callout box',
      '{ type: "callout", title?: string, variant?: "info"|"warning"|"success"|"danger" }',
    ],
    [
      'codeBlock',
      'Syntax-highlighted code',
      '{ code: string, language?: string }',
    ],
    [
      'coverPage',
      'Cover page layout',
      '{ title: string, author?: string, date?: string }',
    ],
    [
      'dataTable',
      'Formatted data table',
      '{ columns: [...], data: [...], striped?: boolean }',
    ],
    [
      'echarts',
      'ECharts chart embedding',
      '{ option: object, width?: number, height?: number }',
    ],
    [
      'meetingMinutes',
      'Meeting minutes template',
      '{ title?: string, date?: string, attendees?: string[], agenda?: [...], notes?: string[] }',
    ],
    ['pageNumber', 'Page numbering', '{ format?: string }'],
    [
      'propertyTable',
      'Key-value property table',
      '{ items: [{ key: string, value: string }] }',
    ],
    [
      'qrcode',
      'QR code generation',
      '{ text: string, width?: number, height?: number }',
    ],
    [
      'signatureBlock',
      'Signature block layout',
      '{ parties: [{ name: string, role?: string }] }',
    ],
    [
      'timeline',
      'Timeline visualization',
      '{ events: [{ date: string, title: string, description?: string }] }',
    ],
    ['watermark', 'Page watermark text', '{ text: string, opacity?: number }'],
  ]

  for (const [name, desc, opts] of pluginDescriptions) {
    parts.push(`- **${name}**: ${desc} — options: ${opts}`)
  }

  return parts.join('\n')
}

/**
 * Generate OpenAI-compatible function/tool definitions
 * for docx-kit document generation.
 *
 * Produces tool definitions that can be used with OpenAI,
 * Anthropic, or other LLM function calling APIs.
 *
 * @param options - Configuration for which tools to include
 * @returns Array of function/tool definition objects
 */
export function generateToolDefinitions(
  options: GenerateToolDefinitionsOptions = {},
): Array<{
  type: 'function'
  function: {
    description: string
    name: string
    parameters: Record<string, unknown>
  }
}> {
  const tools: Array<{
    type: 'function'
    function: {
      description: string
      name: string
      parameters: Record<string, unknown>
    }
  }> = [
    {
      type: 'function',
      function: {
        name: 'create_document',
        description:
          'Create a new .docx document from a docx-kit JSON schema. The schema defines content nodes, styles, and page configuration.',
        parameters: {
          additionalProperties: true,
          required: ['content'],
          type: 'object',
          description:
            'A docx-kit DocxSchema object defining the document content, styles, and configuration.',
          properties: {
            content: {
              type: 'array',
              description:
                'Ordered array of block nodes (headings, paragraphs, tables, etc.)',
              items: {
                additionalProperties: true,
                properties: { type: { type: 'string' } },
                type: 'object',
              },
            },
            page: {
              description: 'Page configuration (size, margins, orientation)',
              type: 'object',
              properties: {
                margin: { type: 'string' },
                size: { enum: ['A3', 'A4', 'Legal', 'Letter'], type: 'string' },
                orientation: {
                  enum: ['landscape', 'portrait'],
                  type: 'string',
                },
              },
            },
            styles: {
              description: 'Named stylesheet entries (className → style rule)',
              type: 'object',
              additionalProperties: {
                additionalProperties: true,
                type: 'object',
              },
            },
          },
        },
      },
    },
    {
      type: 'function',
      function: {
        name: 'validate_schema',
        description:
          'Validate a docx-kit JSON schema. Checks node types, required fields, and plugin references.',
        parameters: {
          required: ['schema'],
          type: 'object',
          properties: {
            schema: {
              additionalProperties: true,
              description: 'The DocxSchema to validate',
              type: 'object',
            },
          },
        },
      },
    },
  ]

  // List plugins tool
  if (options.plugins && options.plugins.length > 0) {
    tools.push({
      type: 'function',
      function: {
        name: 'list_plugins',
        description:
          'List available docx-kit plugins and their option schemas.',
        parameters: {
          type: 'object',
          properties: {
            filter: {
              description: 'Optional filter by plugin name pattern',
              type: 'string',
            },
          },
        },
      },
    })
  }

  // Template-based tools
  if (options.template) {
    const template = BUILTIN_TEMPLATES.find(t => t.name === options.template)
    if (template) {
      tools.push(...generateTemplateTools(template))
    }
  }

  // Always include list_templates
  tools.push({
    type: 'function',
    function: {
      name: 'list_templates',
      description:
        'List all available docx-kit AI templates for document generation.',
      parameters: {
        properties: {},
        type: 'object',
      },
    },
  })

  return tools
}

/**
 * Generate tool definitions for a specific template.
 */
function generateTemplateTools(template: AiTemplate): Array<{
  type: 'function'
  function: {
    description: string
    name: string
    parameters: Record<string, unknown>
  }
}> {
  return [
    {
      type: 'function',
      function: {
        description: `Generate a ${template.name} document using the docx-kit AI template. ${template.description}`,
        name: `apply_template_${template.name}`,
        parameters: {
          ...(template.schema as unknown as Record<string, unknown>),
          required: Object.entries(template.schema.properties)
            .filter(([, def]) => def.required)
            .map(([key]) => key),
        },
      },
    },
  ]
}
