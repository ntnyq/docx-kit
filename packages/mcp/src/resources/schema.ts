/**
 * MCP resource: DocxSchema JSON Schema.
 *
 * Provides a complete JSON Schema for the DocxSchema type,
 * enabling LLMs to understand the expected document structure.
 *
 * @module mcp-server/resources/schema
 */

import {
  buildBlockNodeJsonSchemas,
  buildNestedNodeJsonSchemas,
} from '../schema/blockNodes'

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
      ...buildNestedNodeJsonSchemas(),
      blockNode: {
        oneOf: buildBlockNodeJsonSchemas(),
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
