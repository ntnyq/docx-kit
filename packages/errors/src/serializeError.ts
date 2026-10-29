import { DocxKitError } from './DocxKitError'
import type { DiagnosticCode } from './catalog'

/**
 * Allowlisted product DTO. Never includes stack, cause, or document data.
 */
export interface SerializedDocxKitError {
  message: string
  code?: string
  diagnosticCode?: DiagnosticCode
  docs?: string
  fix?: string
}

/**
 * Serializes diagnostics without exposing cause, stack, or URL credentials.
 * URL userinfo, query strings and fragments are omitted only from the DTO;
 * the original error remains unchanged. Other free-form content is preserved.
 */
export function serializeDocxKitError(error: unknown): SerializedDocxKitError {
  try {
    if (error instanceof DocxKitError) {
      return {
        code: redactDiagnosticUrls(error.code),
        diagnosticCode: error.diagnosticCode,
        message: redactDiagnosticUrls(error.message),
        docs:
          error.docs === undefined
            ? undefined
            : redactDiagnosticUrls(error.docs),
        fix:
          error.fix === undefined ? undefined : redactDiagnosticUrls(error.fix),
      }
    }
    return {
      message: redactDiagnosticUrls(
        error instanceof Error ? error.message : String(error),
      ),
    }
  } catch {
    // JavaScript can throw objects without primitive conversion, or proxies
    // and accessors that throw. Reporting must not replace the original failure.
    return {
      message: 'Unknown error',
    }
  }
}

/**
 * Keeps the URL origin and path useful without guessing which query keys
 * contain credentials. Split adjacent URLs only before a query or fragment;
 * nested URLs inside either belong to the outer URL and must be removed too.
 * Malformed URL-shaped tokens are omitted entirely.
 */
function redactDiagnosticUrls(text: string): string {
  return text.replace(
    /\b(?:https?|wss?|ftp|file):\/\/(?:(?!\b(?:https?|wss?|ftp|file):\/\/)[^\s?#])*(?:[?#]\S*)?/giu,
    (candidate, offset: number) => {
      let end = candidate.length
      // Only unmatched closing brackets can belong to an enclosing array;
      // balanced brackets may be part of an IPv6 host or URL path.
      let unmatchedClosingBrackets = 0
      for (const char of candidate) {
        if (char === '[') {
          unmatchedClosingBrackets--
        }
        if (char === ']') {
          unmatchedClosingBrackets++
        }
      }
      // Preserve angle-link wrappers without treating raw angle brackets in
      // credentials or query values as boundaries that expose their contents.
      while (
        end > 0
        && ('),.;!?"\'`'.includes(candidate[end - 1])
          || (candidate[end - 1] === ']' && unmatchedClosingBrackets > 0)
          || (candidate[end - 1] === '>' && text[offset - 1] === '<'))
      ) {
        if (candidate[end - 1] === ']') {
          unmatchedClosingBrackets--
        }
        end--
      }
      const suffix = candidate.slice(end)
      const value = candidate.slice(0, end)
      try {
        const url = new URL(value)
        if (!url.username && !url.password && !url.search && !url.hash) {
          return candidate
        }
        url.username = ''
        url.password = ''
        url.search = ''
        url.hash = ''
        return `${url.href}${suffix}`
      } catch {
        return `[REDACTED URL]${suffix}`
      }
    },
  )
}
