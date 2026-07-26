/**
 * Node.js platform barrel — re-exports all Node.js-specific APIs.
 *
 * @module node
 */

// Streaming and file-system output
export { saveDocument, streamDocument } from './fs'

// Base64 decoding (via Buffer)
export { dataUrlToUint8Array } from '@docxkit/core'
