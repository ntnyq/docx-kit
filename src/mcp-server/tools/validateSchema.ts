/**
 * MCP tool: validate a docx-kit JSON schema.
 *
 * @module mcp-server/tools/validateSchema
 */

/**
 * Output from the validate_schema MCP tool.
 */
export interface ValidateSchemaOutput {
  /** List of validation errors (empty if valid). */
  errors: ValidationError[]
  /** Whether the schema is valid. */
  valid: boolean
}

/**
 * Validation error detail.
 */
export interface ValidationError {
  /** Error description. */
  message: string
  /** JSON path to the invalid field. */
  path: string
}

/**
 * MCP tool definition for `validate_schema`.
 *
 * Validates a docx-kit JSON schema and returns
 * detailed error information.
 */
export const validateSchemaToolDefinition = {
  name: 'validate_schema',
  description:
    'Validate a docx-kit JSON schema. Checks node types, required fields, and structural correctness.',
  inputSchema: {
    required: ['schema'],
    type: 'object',
    properties: {
      schema: {
        description: 'The DocxSchema JSON object to validate',
        type: 'object',
      },
    },
  },
}

/**
 * Validate a DocxSchema object.
 *
 * Checks for:
 * - Required `content` array
 * - Valid node types in content
 * - Required fields per node type
 *
 * @param schema - — The schema to validate
 * @returns Validation result with errors
 */
export function validateSchema(schema: unknown): ValidateSchemaOutput {
  const errors: ValidationError[] = []
  const doc = schema as Record<string, unknown> | undefined
  const content = doc?.content

  if (!content) {
    errors.push({
      message: 'Required field "content" is missing',
      path: '/content',
    })
  } else if (Array.isArray(content)) {
    // Validate each node
    const validNodeTypes = new Set([
      'bulletList',
      'heading',
      'hyperlink',
      'image',
      'numberedList',
      'pageBreak',
      'paragraph',
      'plugin',
      'sectionBreak',
      'table',
    ])

    for (const [i, element] of content.entries()) {
      const node = element as any
      if (typeof node !== 'object' || node === null) {
        errors.push({
          message: `Node at index ${i} must be an object`,
          path: `/content/${i}`,
        })
        continue
      }

      const nodeType = node.type as string | undefined
      if (!nodeType) {
        errors.push({
          message: `Node at index ${i} missing required "type" field`,
          path: `/content/${i}/type`,
        })
        continue
      }

      if (!validNodeTypes.has(nodeType)) {
        errors.push({
          message: `Invalid node type "${nodeType}" at index ${i}`,
          path: `/content/${i}/type`,
        })
      }

      // Type-specific required fields
      if (nodeType === 'heading' && !node.text) {
        errors.push({
          message: `Heading node at index ${i} missing required "text" field`,
          path: `/content/${i}/text`,
        })
      }
      if (nodeType === 'heading' && !node.level) {
        errors.push({
          message: `Heading node at index ${i} missing required "level" field`,
          path: `/content/${i}/level`,
        })
      }
      if (nodeType === 'paragraph' && !node.text && !node.children) {
        errors.push({
          message: `Paragraph node at index ${i} must have "text" or "children"`,
          path: `/content/${i}`,
        })
      }
      if (nodeType === 'plugin' && !node.name) {
        errors.push({
          message: `Plugin node at index ${i} missing required "name" field`,
          path: `/content/${i}/name`,
        })
      }
      if (nodeType === 'table' && !node.columns) {
        errors.push({
          message: `Table node at index ${i} missing required "columns" field`,
          path: `/content/${i}/columns`,
        })
      }
      if (nodeType === 'table' && !node.data) {
        errors.push({
          message: `Table node at index ${i} missing required "data" field`,
          path: `/content/${i}/data`,
        })
      }
    }
  } else {
    errors.push({
      message: 'Field "content" must be an array',
      path: '/content',
    })
  }

  return {
    errors,
    valid: errors.length === 0,
  }
}
