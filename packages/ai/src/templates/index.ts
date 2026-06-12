/**
 * Built-in AI template barrel export.
 *
 * @module ai/templates
 */

import { invoiceTemplate } from './invoice'
import { letterTemplate } from './letter'
import { reportTemplate } from './report'
import { resumeTemplate } from './resume'
import type { AiTemplate } from '../types'
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

/** All built-in templates. */
export const BUILTIN_TEMPLATES: AiTemplate<any>[] = [
  invoiceTemplate,
  letterTemplate,
  reportTemplate,
  resumeTemplate,
]

/** Names of all built-in templates. */
export const TEMPLATE_LIST = BUILTIN_TEMPLATES.map(t => t.name) as string[]
