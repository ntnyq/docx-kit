/**
 * MCP tool: list and apply AI templates.
 *
 * @module mcp-server/tools/listTemplates
 */

import { BUILTIN_TEMPLATES } from '@docxkit/ai'
import type { AiTemplateInfo } from '@docxkit/ai'

/**
 * MCP tool definition for `list_templates`.
 *
 * Returns all available docx-kit AI templates.
 */
export const listTemplatesToolDefinition = {
  name: 'list_templates',
  description:
    'List all available docx-kit AI templates for document generation.',
  inputSchema: {
    properties: {},
    type: 'object',
  },
}

/**
 * Build template info list from built-in templates.
 *
 * @returns Array of template info (without generator)
 */
export function buildTemplateInfoList(): AiTemplateInfo[] {
  return BUILTIN_TEMPLATES.map(t => ({
    description: t.description,
    name: t.name,
    schema: t.schema,
    systemPrompt: t.systemPrompt,
  }))
}
