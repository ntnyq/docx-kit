/**
 * MCP tool: generate a document from a template + data.
 *
 * @module mcp-server/tools/applyTemplate
 */

import { BUILTIN_TEMPLATES, validateTemplateData } from '@docxkit/ai'
import type { DocxSchema } from '@docxkit/core'

/**
 * Input for the apply_template MCP tool.
 */
export interface ApplyTemplateInput {
  /**
   * Template parameter values (match template schema).
   */
  data: Record<string, unknown>
  /**
   * Name of the template to apply.
   */
  template: string
}

/**
 * Output from the apply_template MCP tool.
 */
export interface ApplyTemplateOutput {
  /**
   * The generated DocxSchema.
   */
  schema: DocxSchema
  /**
   * Name of the template used.
   */
  templateName: string
}

/**
 * MCP tool definition for `apply_template`.
 *
 * Generates a DocxSchema from a named template + data params.
 */
export const applyTemplateToolDefinition = {
  name: 'apply_template',
  description:
    'Generate a docx-kit document schema from an AI template and provided data parameters.',
  inputSchema: {
    required: ['template', 'data'],
    type: 'object',
    properties: {
      data: {
        description: 'Template parameter values (must match template schema)',
        type: 'object',
      },
      template: {
        type: 'string',
        description:
          'Name of the template to apply (report, invoice, resume, letter)',
      },
    },
  },
}

/**
 * Apply a template with given data to produce a DocxSchema.
 *
 * @param templateName - — Template name (e.g. 'report')
 * @param data - — Template parameters
 * @returns Generated DocxSchema and template name, or null if template not found
 * @throws {TypeError} If parameters do not match the template, with JSON Pointer paths
 */
export function applyTemplate(
  templateName: string,
  data: Record<string, unknown>,
): ApplyTemplateOutput | null {
  const template = BUILTIN_TEMPLATES.find(t => t.name === templateName)
  if (!template) {
    return null
  }

  const errors = validateTemplateData(template.schema, data)
  if (errors.length > 0) {
    throw new TypeError(
      `Invalid parameters for template "${templateName}":\n${errors.map(error => `${error.path || '/'}: ${error.message}`).join('\n')}`,
    )
  }

  // The selected template's heterogeneous parameter type is checked above.
  const schema = template.generate(data as never)
  return {
    schema,
    templateName,
  }
}
