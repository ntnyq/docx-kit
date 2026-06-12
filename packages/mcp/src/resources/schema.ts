/**
 * MCP resource: DocxSchema JSON Schema.
 *
 * Provides a complete JSON Schema for the DocxSchema type,
 * enabling LLMs to understand the expected document structure.
 *
 * @module mcp-server/resources/schema
 */

/**
 * DocxSchema JSON Schema for MCP resource exposure.
 *
 * This schema describes the expected structure of a
 * docx-kit document schema, suitable for LLM reference.
 */
export const docxSchemaResource = {
  mimeType: 'application/json',
  name: 'docx-kit://schema',
  description:
    'The complete docx-kit DocxSchema JSON Schema. Describes all node types, their required fields, and how they combine to form a document.',
  schema: {
    $id: 'docx-kit://schema',
    $schema: 'http://json-schema.org/draft-07/schema#',
    description: 'A docx-kit document schema',
    required: ['content'],
    title: 'DocxSchema',
    type: 'object',
    definitions: {
      blockNode: {
        oneOf: [
          {
            description: 'A bullet list',
            required: ['type', 'items'],
            type: 'object',
            properties: {
              type: { const: 'bulletList' },
              items: {
                items: { type: 'string' },
                type: 'array',
              },
            },
          },
          {
            description: 'A heading (h1–h6)',
            required: ['type', 'level', 'text'],
            type: 'object',
            properties: {
              className: { type: 'string' },
              level: { enum: [1, 2, 3, 4, 5, 6], type: 'number' },
              style: { additionalProperties: true, type: 'object' },
              text: { type: 'string' },
              type: { const: 'heading' },
            },
          },
          {
            description: 'A hyperlink',
            required: ['type', 'url', 'children'],
            type: 'object',
            properties: {
              type: { const: 'hyperlink' },
              url: { type: 'string' },
              children: {
                items: { type: 'string' },
                type: 'array',
              },
            },
          },
          {
            description: 'An image',
            required: ['type', 'data'],
            type: 'object',
            properties: {
              alt: { type: 'string' },
              data: { type: 'string' },
              height: { type: 'number' },
              type: { const: 'image' },
              width: { type: 'number' },
              imageType: {
                enum: ['png', 'jpeg', 'jpg', 'gif', 'bmp'],
                type: 'string',
              },
            },
          },
          {
            description: 'A numbered list',
            required: ['type', 'items'],
            type: 'object',
            properties: {
              type: { const: 'numberedList' },
              items: {
                items: { type: 'string' },
                type: 'array',
              },
            },
          },
          {
            description: 'A page break',
            properties: { type: { const: 'pageBreak' } },
            required: ['type'],
            type: 'object',
          },
          {
            description: 'A paragraph of text',
            required: ['type'],
            type: 'object',
            properties: {
              className: { type: 'string' },
              style: { additionalProperties: true, type: 'object' },
              text: { type: 'string' },
              type: { const: 'paragraph' },
              children: {
                items: { additionalProperties: true, type: 'object' },
                type: 'array',
              },
            },
          },
          {
            description: 'A plugin invocation',
            required: ['type', 'name', 'options'],
            type: 'object',
            properties: {
              name: { type: 'string' },
              options: { additionalProperties: true, type: 'object' },
              type: { const: 'plugin' },
            },
          },
          {
            description: 'A section break',
            required: ['type'],
            type: 'object',
            properties: {
              config: { additionalProperties: true, type: 'object' },
              type: { const: 'sectionBreak' },
            },
          },
          {
            description: 'A data table',
            required: ['type', 'columns', 'data'],
            type: 'object',
            properties: {
              header: { type: 'boolean' },
              striped: { type: 'boolean' },
              type: { const: 'table' },
              columns: {
                type: 'array',
                items: {
                  required: ['key', 'title'],
                  type: 'object',
                  properties: {
                    key: { type: 'string' },
                    title: { type: 'string' },
                    width: { type: 'string' },
                    align: {
                      enum: ['left', 'center', 'right'],
                      type: 'string',
                    },
                  },
                },
              },
              data: {
                items: { additionalProperties: true, type: 'object' },
                type: 'array',
              },
            },
          },
        ],
      },
      styleRule: {
        additionalProperties: true,
        description: 'CSS-like style properties for a node',
        type: 'object',
        properties: {
          backgroundColor: { type: 'string' },
          borderBottom: { type: 'string' },
          borderLeft: { type: 'string' },
          borderRight: { type: 'string' },
          borderTop: { type: 'string' },
          color: { type: 'string' },
          fontFamily: { type: 'string' },
          fontSize: { type: 'number' },
          fontWeight: { enum: ['bold', 'normal', 'semibold'], type: 'string' },
          lineHeight: { type: 'number' },
          marginBottom: { type: 'number' },
          marginLeft: { type: 'number' },
          marginTop: { type: 'number' },
          textAlign: {
            enum: ['center', 'justify', 'left', 'right'],
            type: 'string',
          },
        },
      },
    },
    properties: {
      content: {
        description: 'Ordered array of block nodes',
        items: { $ref: '#/definitions/blockNode' },
        type: 'array',
      },
      page: {
        description: 'Page configuration',
        type: 'object',
        properties: {
          margin: { type: 'string' },
          orientation: { enum: ['landscape', 'portrait'], type: 'string' },
          size: { enum: ['A3', 'A4', 'Legal', 'Letter'], type: 'string' },
        },
      },
      plugins: {
        description: 'Plugin sources to load before rendering',
        type: 'array',
        items: {
          required: ['type'],
          type: 'object',
          properties: {
            plugin: { description: 'Inline plugin instance', type: 'object' },
            type: { enum: ['inline', 'npm', 'url', 'local'], type: 'string' },
          },
        },
      },
      styles: {
        additionalProperties: { $ref: '#/definitions/styleRule' },
        description: 'Named stylesheet entries',
        type: 'object',
      },
    },
  },
}
