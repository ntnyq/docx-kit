/**
 * AI module entry point.
 *
 * Provides templates, prompt builder, schema guide,
 * and tool definition generation for LLM integration.
 *
 * @module ai
 */

export { BUILTIN_PLUGIN_GUIDES } from './plugins'
export { buildFreeformPrompt, buildPrompt } from './promptBuilder'
export { STYLE_RULE_SCHEMA, UNIT_VALUE_SCHEMA } from './styleSchema'
export { generateSchemaGuide, generateToolDefinitions } from './schemaGuide'
export { toTemplateJsonSchema, validateTemplateData } from './templateSchema'
export {
  referenceStyleSchemas,
  STYLE_SCHEMA_DEFINITIONS,
} from './schemaDefinitions'
export {
  BUILTIN_PLUGIN_OPTION_SCHEMAS,
  BUILTIN_PLUGIN_SCHEMA_CONDITIONS,
  SECTION_BREAK_SCHEMA_CONDITION,
} from './pluginSchema'
export {
  BUILTIN_TEMPLATES,
  invoiceTemplate,
  letterTemplate,
  reportTemplate,
  resumeTemplate,
  TEMPLATE_LIST,
} from './templates'
export type { LetterParams } from './templates/letter'
export type { ReportParams } from './templates/report'
export type { ResumeParams } from './templates/resume'
export type { InvoiceParams } from './templates/invoice'

export type { TemplateValidationError } from './templateSchema'
export type {
  AiTemplate,
  AiTemplateInfo,
  AiTemplateSchema,
  GenerateToolDefinitionsOptions,
} from './types'
