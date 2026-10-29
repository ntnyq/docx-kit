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
import { STYLE_RULE_SCHEMA, UNIT_VALUE_SCHEMA } from '../schema/styles'

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
      styleRule: STYLE_RULE_SCHEMA,
      blockNode: {
        oneOf: buildBlockNodeJsonSchemas(),
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
          margin: STYLE_RULE_SCHEMA.properties.margin,
          orientation: { enum: ['landscape', 'portrait'], type: 'string' },
          size: {
            anyOf: [
              { enum: ['A3', 'A4', 'Legal', 'Letter'] },
              {
                required: ['height', 'width'],
                type: 'object',
                properties: {
                  height: UNIT_VALUE_SCHEMA,
                  width: UNIT_VALUE_SCHEMA,
                },
              },
            ],
          },
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
