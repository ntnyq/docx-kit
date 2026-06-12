/**
 * Node.js platform barrel — re-exports all Node.js-specific APIs.
 *
 * @module node
 */

// File system save
export { saveDocument } from './fs'

// Base64 decoding (via Buffer)
export { dataUrlToUint8Array } from '@docxkit/core'
