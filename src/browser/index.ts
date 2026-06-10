/**
 * Browser platform barrel — re-exports all browser-specific APIs.
 *
 * @module browser
 */

// Blob / image data normalization
export { dataUrlToUint8Array, normalizeImageData } from './dom'

// Base64 decoding (via atob)
