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
import { PAGE_SCHEMA, PLUGIN_SOURCES_SCHEMA } from '../schema/document'
import { STYLE_SCHEMA_DEFINITIONS } from '../schema/styles'

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
      ...STYLE_SCHEMA_DEFINITIONS,
      ...buildNestedNodeJsonSchemas(),
      styleRule: STYLE_SCHEMA_DEFINITIONS.styleRule,
      blockNode: {
        oneOf: buildBlockNodeJsonSchemas(),
      },
    },
    properties: {
      page: { ...PAGE_SCHEMA, description: 'Page configuration' },
      content: {
        description: 'Ordered array of block nodes',
        items: { $ref: '#/definitions/blockNode' },
        type: 'array',
      },
      plugins: {
        ...PLUGIN_SOURCES_SCHEMA,
        description: 'Plugin sources to load before rendering',
      },
      styles: {
        additionalProperties: { $ref: '#/definitions/styleRule' },
        description: 'Named stylesheet entries',
        type: 'object',
      },
    },
  },
}
