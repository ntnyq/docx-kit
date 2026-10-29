import { DIAGNOSTIC_CODES, diagnostics } from './catalog'
import { DocxKitError } from './DocxKitError'
import type { Diagnostic } from 'nostics'
import type { ErrorCode, ErrorParams } from './catalog'

/**
 * Runtime context for a diagnostic. The cause is preserved by reference.
 */
export interface CreateDocxKitErrorOptions {
  cause?: unknown
  /**
   * Retains an existing message when adapting a legacy call site.
   */
  message?: string
}

/**
 * Creates a silent, typed diagnostic while preserving DocxKitError identity.
 */
export function createDocxKitError<Code extends ErrorCode>(
  code: Code,
  params: ErrorParams<NoInfer<Code>>,
  options: CreateDocxKitErrorOptions = {},
): DocxKitError {
  // TypeScript loses the correlation between a generic key and its handle.
  // Each handle's params still come directly from its catalog definition.
  const diagnose = diagnostics[DIAGNOSTIC_CODES[code]] as (
    params: ErrorParams<Code>,
  ) => Diagnostic<unknown>
  const diagnostic = diagnose(params)
  const error = new DocxKitError(
    code,
    options.message ?? diagnostic.message,
    Object.hasOwn(options, 'cause') ? options.cause : params.cause,
    {
      docs: diagnostic.docs,
      fix: diagnostic.fix,
      sources: diagnostic.sources,
    },
  )
  Error.captureStackTrace?.(error, createDocxKitError)
  return error
}
