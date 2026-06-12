/**
 * MCP tool: create a new .docx document from a schema.
 *
 * @module mcp-server/tools/createDocx
 */

import type { DocxSchema } from '../../builder/createDocx'

/**
 * Input schema for the create_document MCP tool.
 */
export interface CreateDocumentInput {
  /** File path for the output .docx file. */
  outputPath: string
  /** The docx-kit DocxSchema defining the document. */
  schema: DocxSchema
}

/**
 * Output from the create_document MCP tool.
 */
export interface CreateDocumentOutput {
  /** Path of the created file. */
  filePath: string
  /** File size in bytes. */
  size: number
}

/**
 * MCP tool definition for `create_document`.
 *
 * Creates a new .docx file from a docx-kit JSON schema.
 */
export const createDocxToolDefinition = {
  name: 'create_document',
  description:
    'Create a new .docx document from a docx-kit JSON schema. The schema defines content nodes, styles, and page configuration.',
  inputSchema: {
    required: ['schema', 'outputPath'],
    type: 'object',
    properties: {
      outputPath: {
        description: 'File path for the output .docx file',
        type: 'string',
      },
      schema: {
        type: 'object',
        description:
          'A docx-kit DocxSchema object with content, styles, and page config',
      },
    },
  },
}
