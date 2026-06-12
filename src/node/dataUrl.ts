/**
 * Node.js base64 data-URL decoder (re-exports the shared implementation).
 *
 * The shared implementation in {@link ../utils/dataUrl} auto-detects
 * the runtime and uses `Buffer.from` in Node.js environments.
 *
 * @module node/dataUrl
 */

export { dataUrlToUint8Array } from '../utils/dataUrl'
