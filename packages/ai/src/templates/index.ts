/**
 * Built-in AI template barrel export.
 *
 * @module ai/templates
 */

import { invoiceTemplate } from './invoice'
import { letterTemplate } from './letter'
import { reportTemplate } from './report'
import { resumeTemplate } from './resume'
import type { InvoiceParams } from './invoice'
import type { LetterParams } from './letter'
import type { ReportParams } from './report'
import type { ResumeParams } from './resume'

export { invoiceTemplate }
export type { InvoiceParams }
export { letterTemplate }
export type { LetterParams }
export { reportTemplate }
export type { ReportParams }
export { resumeTemplate }
export type { ResumeParams }

export type BuiltinAiTemplate =
  | typeof invoiceTemplate
  | typeof letterTemplate
  | typeof reportTemplate
  | typeof resumeTemplate

/**
 * All built-in templates.
 */
export const BUILTIN_TEMPLATES = [
  invoiceTemplate,
  letterTemplate,
  reportTemplate,
  resumeTemplate,
] as const satisfies readonly BuiltinAiTemplate[]

export type BuiltinTemplateName = (typeof BUILTIN_TEMPLATES)[number]['name']

/**
 * Names of all built-in templates.
 */
export const TEMPLATE_LIST: BuiltinTemplateName[] = BUILTIN_TEMPLATES.map(
  template => template.name,
)
