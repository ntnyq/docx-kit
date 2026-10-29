import { DIAGNOSTIC_CODES } from '@docxkit/errors'
import type { SerializedDocxKitError } from '@docxkit/errors'

export { serializeDocxKitError } from '@docxkit/errors'
export type { SerializedDocxKitError } from '@docxkit/errors'

/**
 * Links only published diagnostics to the current documentation language.
 */
export function getDiagnosticDocsUrl(
  error: SerializedDocxKitError,
  locale: string,
): string | undefined {
  if (
    !error.diagnosticCode
    || !Object.values(DIAGNOSTIC_CODES).includes(error.diagnosticCode)
  ) {
    return undefined
  }
  const prefix = locale === 'zh-CN' || locale === 'zh-TW' ? `/${locale}` : ''
  return `https://docx-kit.ntnyq.dev${prefix}/errors/${error.diagnosticCode.toLowerCase()}/`
}
