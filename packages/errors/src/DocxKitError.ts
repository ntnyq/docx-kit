import { DIAGNOSTIC_CODES, diagnosticDefinitions } from './catalog'
import type { DiagnosticCode, ErrorCode } from './catalog'

/**
 * Optional diagnostics attached without changing the legacy constructor.
 */
export interface DocxKitErrorDetails {
  docs?: string
  fix?: string
  sources?: readonly string[]
}

/**
 * Shared error identity, legacy code and original cause are preserved.
 * Known codes gain a full diagnostic code and documentation URL.
 */
export class DocxKitError extends Error {
  readonly cause: unknown
  readonly code: string | ErrorCode
  readonly diagnosticCode?: DiagnosticCode
  readonly docs?: string
  readonly fix?: string
  readonly sources?: readonly string[]

  // eslint-disable-next-line unicorn/custom-error-definition -- Preserve the public code-first constructor.
  constructor(
    code: string | ErrorCode,
    message: string,
    cause?: unknown,
    details?: DocxKitErrorDetails,
  ) {
    super(message)
    this.name = 'DocxKitError'
    this.code = code
    this.cause = cause
    if (Object.hasOwn(DIAGNOSTIC_CODES, code)) {
      const diagnosticCode = DIAGNOSTIC_CODES[code as ErrorCode]
      this.diagnosticCode = diagnosticCode
      this.docs = `https://docx-kit.ntnyq.dev/errors/${diagnosticCode.toLowerCase()}/`
      const fix = diagnosticDefinitions[diagnosticCode].fix
      this.fix =
        typeof fix === 'string'
          ? fix
          : 'Check the plugin source, platform, module exports, and security policy.'
    }
    this.fix = details?.fix ?? this.fix
    this.docs = details?.docs ?? this.docs
    this.sources = details?.sources
  }
}
