/**
 * Build LLM prompts from templates and user input.
 *
 * The prompt builder combines a template's system prompt
 * with user-provided context to create a complete prompt
 * for LLM document generation.
 *
 * @module ai/promptBuilder
 */

import { BUILTIN_PLUGIN_GUIDES } from './plugins'
import { toTemplateJsonSchema } from './templateSchema'
import type { AiTemplate } from './types'

/**
 * Build a prompt for generating a document from scratch
 * (without a template).
 *
 * Provides general docx-kit schema guidance.
 *
 * @param documentType - The type of document to generate
 * @param userContext - Additional requirements from the user
 * @returns A complete prompt string
 */
export function buildFreeformPrompt(
  documentType: string,
  userContext?: string,
): string {
  const parts = [
    `You are a professional document generator. Generate a docx-kit JSON schema for a ${documentType}.`,
    '',
    '## Available Node Types',
    '- heading: { type: "heading", level: 1-6, text: "..." }',
    '- paragraph: { type: "paragraph", text: "..." }',
    '- table: { type: "table", columns: [...], data: [...] }',
    '- bulletList: { type: "bulletList", items: [...] }',
    '- numberedList: { type: "numberedList", items: [...] }',
    '- image: { type: "image", data: "...", width: N, height: N }',
    '- pageBreak: { type: "pageBreak" }',
    '- hyperlink: { type: "hyperlink", url: "...", children: [...] }',
    '- plugin: { type: "plugin", name: "...", options: { ... } }',
    '',
    '## Available Plugins',
    ...BUILTIN_PLUGIN_GUIDES.map(
      plugin => `- ${plugin.name}: ${plugin.description}`,
    ),
  ]

  if (userContext) {
    parts.push('', '## User Requirements', userContext)
  }

  parts.push(
    '',
    '## Output Format',
    'Return a valid docx-kit DocxSchema JSON: { content: [...], styles: { ... }, page: { ... } }',
  )

  return parts.join('\n')
}

/**
 * Build a complete LLM prompt from a template and user context.
 *
 * Combines the template's system prompt with the user's
 * content requirements and schema reference.
 *
 * @param template - The AI template to build a prompt for
 * @param userContext - Additional context from the user
 * @returns A complete prompt string for LLM consumption
 */
export function buildPrompt<TParams extends object>(
  template: AiTemplate<TParams>,
  userContext?: string,
): string {
  const parts = [
    template.systemPrompt,
    '',
    `## Template: ${template.name}`,
    `Description: ${template.description}`,
    '',
    '### Input Schema',
    JSON.stringify(toTemplateJsonSchema(template.schema), null, 2),
  ]

  // User context
  if (userContext) {
    parts.push('', '## User Requirements', userContext)
  }

  // Output format reminder
  parts.push(
    '',
    '## Output Format',
    'Return a valid docx-kit DocxSchema JSON object with a "content" array of nodes.',
    'Each node must have a "type" field (heading, paragraph, table, bulletList, pageBreak, plugin, etc.).',
  )

  return parts.join('\n')
}
