/**
 * AI module entry point.
 *
 * Provides templates, prompt builder, schema guide,
 * and tool definition generation for LLM integration.
 *
 * @module ai
 */

export { buildFreeformPrompt, buildPrompt } from './promptBuilder'
export { generateSchemaGuide, generateToolDefinitions } from './schemaGuide'
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
export type {
  AiTemplate,
  AiTemplateInfo,
  AiTemplateSchema,
  GenerateToolDefinitionsOptions,
} from './types'
