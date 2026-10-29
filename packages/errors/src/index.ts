export { DocxKitError } from './DocxKitError'
export { createDocxKitError } from './createError'
export { serializeDocxKitError } from './serializeError'
/**
 * Shared compatible errors backed by a silent Nostics catalog.
 *
 * @module errors
 */
export { DIAGNOSTIC_CODES, ERROR_CODES } from './catalog'
export type { DocxKitErrorDetails } from './DocxKitError'
export type { CreateDocxKitErrorOptions } from './createError'
export type { SerializedDocxKitError } from './serializeError'
export type {
  DiagnosticCode,
  ErrorCode,
  ErrorParams,
  PluginLoadReason,
} from './catalog'
