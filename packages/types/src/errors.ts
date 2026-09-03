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
  /**
   * Document export failed.
   */
  EXPORT_FAILED: 'EXPORT_FAILED',
  /**
   * Image data is empty, corrupt, or unsupported.
   */
  IMAGE_INVALID_DATA: 'IMAGE_INVALID_DATA',
  /**
   * Plugin manifest JSON is missing required fields or has invalid structure.
   */
  MANIFEST_INVALID: 'MANIFEST_INVALID',
  /**
   * Plugin manifest file was not found at the expected path.
   */
  MANIFEST_MISSING: 'MANIFEST_MISSING',
  /**
   * A required plugin dependency was not found or version is incompatible.
   */
  PLUGIN_DEPENDENCY_MISSING: 'PLUGIN_DEPENDENCY_MISSING',
  /**
   * Plugin failed to load (import error, network error, etc.).
   */
  PLUGIN_LOAD_FAILED: 'PLUGIN_LOAD_FAILED',
  /**
   * A plugin node referenced an unregistered plugin.
   */
  PLUGIN_NOT_REGISTERED: 'PLUGIN_NOT_REGISTERED',

  // ---- Plugin ecosystem errors ----

  /**
   * A plugin's `render()` method threw an error.
   */
  PLUGIN_RENDER_FAILED: 'PLUGIN_RENDER_FAILED',
  /**
   * Plugin's `docxKit` semver range is incompatible with the running version.
   */
  PLUGIN_VERSION_MISMATCH: 'PLUGIN_VERSION_MISMATCH',
  /**
   * A `className` referenced a stylesheet key that doesn't exist.
   */
  STYLE_UNKNOWN_CLASS: 'STYLE_UNKNOWN_CLASS',
  /**
   * Table was created with no columns.
   */
  TABLE_INVALID_COLUMNS: 'TABLE_INVALID_COLUMNS',
  /**
   * Encountered a node type that has no registered compiler.
   */
  UNKNOWN_NODE_TYPE: 'UNKNOWN_NODE_TYPE',
} as const

/**
 * Union type of all known error codes.
 */
export type ErrorCode = (typeof ERROR_CODES)[keyof typeof ERROR_CODES]

/**
 * Structured error class for docx-kit.
 *
 * Carries a machine-readable `code`, a human-readable `message`,
 * and optionally the underlying `cause`.
 */
export class DocxKitError extends Error {
  /**
   * The underlying error that caused this failure (if any).
   */
  readonly cause: unknown
  /**
   * Machine-readable error code.
   */
  readonly code: string | ErrorCode

  /**
   * @param code - — Known error code from {@link ERROR_CODES}
   * @param message - — Human-readable description
   * @param cause - — Optional underlying error
   */
  // eslint-disable-next-line unicorn/custom-error-definition -- Public API uses a code-first signature for structured errors.
  constructor(code: string | ErrorCode, message: string, cause?: unknown) {
    super(message)
    this.name = 'DocxKitError'
    this.code = code
    this.cause = cause
  }
}
