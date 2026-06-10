/**
 * Error handling types and classes for docx-kit.
 *
 * @module errors
 */

/**
 * Well-known error codes used throughout the library.
 *
 * Consumers can match against these codes for structured error handling.
 */
export const ERROR_CODES = {
  /** Document export failed. */
  EXPORT_FAILED: 'EXPORT_FAILED',
  /** Image data is empty, corrupt, or unsupported. */
  IMAGE_INVALID_DATA: 'IMAGE_INVALID_DATA',
  /** Could not determine image format. */
  IMAGE_UNKNOWN_TYPE: 'IMAGE_UNKNOWN_TYPE',
  /** A plugin node referenced an unregistered plugin. */
  PLUGIN_NOT_REGISTERED: 'PLUGIN_NOT_REGISTERED',
  /** A plugin's `render()` method threw an error. */
  PLUGIN_RENDER_FAILED: 'PLUGIN_RENDER_FAILED',
  /** Invalid unit string was provided. */
  STYLE_INVALID_UNIT: 'STYLE_INVALID_UNIT',
  /** A `className` referenced a stylesheet key that doesn't exist. */
  STYLE_UNKNOWN_CLASS: 'STYLE_UNKNOWN_CLASS',
  /** Table was created with no columns. */
  TABLE_INVALID_COLUMNS: 'TABLE_INVALID_COLUMNS',
  /** Encountered a node type that has no registered compiler. */
  UNKNOWN_NODE_TYPE: 'UNKNOWN_NODE_TYPE',
} as const

/** Union type of all known error codes. */
export type ErrorCode = (typeof ERROR_CODES)[keyof typeof ERROR_CODES]

/**
 * Structured error class for docx-kit.
 *
 * Carries a machine-readable `code`, a human-readable `message`,
 * and optionally the underlying `cause`.
 *
 * @example
 * ```ts
 * try {
 *   await doc.save('output.docx')
 * } catch (err) {
 *   if (err instanceof DocxKitError && err.code === ERROR_CODES.EXPORT_FAILED) {
 *     console.error('Export failed:', err.message)
 *   }
 * }
 * ```
 */
export class DocxKitError extends Error {
  /** The underlying error that caused this failure (if any). */
  readonly cause: unknown
  /** Machine-readable error code. */
  readonly code: string | ErrorCode

  /**
   * @param code - — Known error code from {@link ERROR_CODES}
   * @param message - — Human-readable description
   * @param cause - — Optional underlying error
   */
  constructor(code: string | ErrorCode, message: string, cause?: unknown) {
    super(message)
    this.name = 'DocxKitError'
    this.code = code
    this.cause = cause
  }
}
